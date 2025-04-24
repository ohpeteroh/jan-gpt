import os
import tempfile
from pathlib import Path
import streamlit as st
import pandas as pd
import subprocess
from pptx import Presentation
import pyxlsb

from langchain.docstore.document import Document
from langchain_community.vectorstores import FAISS
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain_community.document_loaders import TextLoader, PyPDFLoader, UnstructuredWordDocumentLoader
from langchain_community.tools import DuckDuckGoSearchRun

# 🔐 API 및 초기 설정
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", "sk-...")
embedding = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)
llm = ChatOpenAI(model="gpt-4", temperature=0.2, openai_api_key=OPENAI_API_KEY)
search_tool = DuckDuckGoSearchRun()
splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=100)
DB_PATH = "faiss_index"

# 📄 다양한 파일 로딩
def load_and_split_file(tmp_path, suffix):
    docs = []
    if suffix == ".txt":
        with open(tmp_path, encoding="utf-8") as f:
            text = f.read()
        docs = [Document(page_content=text)]
    elif suffix == ".pdf":
        docs = PyPDFLoader(tmp_path).load()
    elif suffix == ".docx":
        docs = UnstructuredWordDocumentLoader(tmp_path).load()
    elif suffix == ".pptx":
        prs = Presentation(tmp_path)
        text = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
        docs = [Document(page_content=text)]
    elif suffix == ".hwp":
        result = subprocess.run(['hwp5txt', tmp_path], stdout=subprocess.PIPE, encoding='utf-8')
        docs = [Document(page_content=result.stdout)]
    elif suffix in [".xlsx", ".xlsm"]:
        df = pd.read_excel(tmp_path, engine='openpyxl')
        docs = [Document(page_content=df.to_string())]
    elif suffix == ".xlsb":
        with pyxlsb.open_workbook(tmp_path) as wb:
            sheet = wb.get_sheet(1)
            data = "\n".join(["\t".join([str(cell.v) for cell in row]) for row in sheet.rows()])
        docs = [Document(page_content=data)]
    elif suffix == ".csv":
        df = pd.read_csv(tmp_path)
        docs = [Document(page_content=df.to_string())]

    chunks = splitter.split_documents(docs)
    if not os.path.exists(DB_PATH):
        db = FAISS.from_documents(chunks, embedding)
    else:
        db = FAISS.load_local(DB_PATH, embedding)
        db.add_documents(chunks)
    db.save_local(DB_PATH)
    return True

# 🚀 Streamlit 인터페이스
st.set_page_config(page_title="Jan GPT", layout="wide")
st.title("📂 Jan GPT - 파일 기반 검색 & 리서치 AI")

uploaded_file = st.file_uploader(
    "📤 파일 업로드 (.txt, .pdf, .docx, .pptx, .hwp, .xlsx, .xlsm, .xlsb, .csv)",
    type=["txt", "pdf", "docx", "pptx", "hwp", "xlsx", "xlsm", "xlsb", "csv"]
)

if uploaded_file:
    suffix = Path(uploaded_file.name).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    success = load_and_split_file(tmp_path, suffix)
    if success:
        st.success("✅ 파일이 자동으로 벡터화되어 학습되었습니다.")

query = st.text_input("❓ 질문을 입력하세요")
search_mode = st.radio("검색 모드", ["일반 검색", "심층 리서치"], horizontal=True)

if query:
    if os.path.exists(DB_PATH):
        db = FAISS.load_local(DB_PATH, embedding)
        docs = db.similarity_search(query, k=5)
        doc_context = "\n\n".join([doc.page_content for doc in docs])
    else:
        doc_context = "(문서가 없습니다)"
    web_results = search_tool.run(query)

    if search_mode == "심층 리서치":
        prompt = f"""
[문서 기반 정보]
{doc_context}

[웹 검색 정보]
{web_results}

이 정보를 바탕으로 질문 '{query}'에 대해 다음 항목을 포함하여 심층 분석 보고서를 작성해주세요:
1. 핵심 요약
2. 주요 근거 및 배경 정보
3. 전략적 시사점 및 조언
"""
    else:
        prompt = f"""
[문서 기반 정보]
{doc_context}

[웹 검색 정보]
{web_results}

위 정보를 바탕으로 '{query}'에 답변해 주세요.
"""
    response = llm.invoke(prompt)
    st.markdown("### 💬 GPT 응답")
    st.write(response.content)
