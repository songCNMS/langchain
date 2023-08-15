from langchain import OpenAI, PromptTemplate, LLMChain
from langchain.text_splitter import CharacterTextSplitter,NLTKTextSplitter,TextSplitter
from langchain.chains.mapreduce import MapReduceChain
from langchain.prompts import PromptTemplate
from langchain.chat_models import AzureChatOpenAI
import os
from typing import Any, List
import openai
import openai_config
from langchain.chains.summarize import load_summarize_chain
import nltk
from langchain.docstore.document import Document
import fitz
import re
from docx2python import docx2python
import chinese_converter
# import win32com.client as win32
# from win32com.client import constants
# nltk.download('punkt')


def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.visible = 0
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)


def list_files_recursive(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


def preprocess(text):
    text = text.strip().replace("\n", "")
    text = text.replace("\t", "")
    text = re.sub("\s+", "", text)
    return text


def pdf_to_text(path, start_page=1, end_page=None):
    doc = fitz.open(path)
    total_pages = doc.page_count
    if end_page is None:
        end_page = total_pages
    text_list = []
    for i in range(start_page - 1, end_page):
        text = chinese_converter.to_simplified(doc.load_page(i).get_text("text"))
        text = preprocess(text)
        text_list.append(text)
    doc.close()
    return "".join(text_list)


def doc_to_text(path):
    doc_result = docx2python(path)
    text = preprocess(chinese_converter.to_simplified(doc_result.text))
    return text


def get_text(file_loc):
    ext = os.path.splitext(file_loc)[1]
    if ext.lower() == ".pdf":
        text = pdf_to_text(file_loc)
    elif ext.lower() in [".docx"]:
        text = doc_to_text(file_loc)
    else:
        raise Exception(f"Unsupported file types: {ext.lower()}")
    return text


def get_chunks(text, max_len_per_chunk=2000):
    if len(text) <= max_len_per_chunk: return [text]
    paragraph_list = text.split("。")
    chunk_list = []
    start = 0
    while start < len(paragraph_list):
        end = start
        while end < len(paragraph_list):
            end += 1
            chunk = "。".join(paragraph_list[start:end])
            if len(chunk) > max_len_per_chunk:
                chunk_list.append(chunk)
                break
        if end == len(paragraph_list): break
        start = max(start+1, end-2)
    return chunk_list


class ChineseSplitter(TextSplitter):
    def __init__(self, **kwargs: Any):
        """Create a new TextSplitter."""
        super().__init__(**kwargs)

    def split_text(self, text: str) -> List[str]:
        return get_chunks(text)



if __name__ == "__main__":
    llm = AzureChatOpenAI(temperature=0, deployment_name="gpt-4-32k")
    text_splitter = ChineseSplitter()
    file_dir = "./data/食品政策文档/"
    abstract_dir =f"{file_dir}/abstracts/"
    os.makedirs(abstract_dir, exist_ok=True)
    product_name = "荔枝"
    for file_loc in list_files_recursive(file_dir):
        print(file_loc)
        ext = os.path.splitext(file_loc)[1]
        if ext.lower() not in [".pdf", ".docx"]: continue
        text = get_text(file_loc)
        file_name = os.path.basename(file_loc)
        texts = text_splitter.split_text(text)
        print("size of texts: ", len(texts))
        docs = [Document(page_content=t) for t in texts]
        prompt_template = f"请从如下描述中提取并总结出跟{product_name}相关的进出口和物品寄递政策。\n"
        prompt_template += """
        {text}
        摘要为:"""
        PROMPT = PromptTemplate(template=prompt_template, input_variables=["text"])
        chain = load_summarize_chain(llm, chain_type="map_reduce", return_intermediate_steps=True, map_prompt=PROMPT, combine_prompt=PROMPT)
        abstract = chain({"input_documents": docs}, return_only_outputs=True)
        print(abstract)
        with open(f"{abstract_dir}/{file_name}_{product_name}_abstract.txt", "w") as f:
            f.writelines(abstract["output_text"])



