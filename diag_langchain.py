import sys
try:
    from langchain.chains import ConversationalRetrievalChain
    print("FOUND in langchain.chains")
except ImportError:
    print("NOT FOUND in langchain.chains")

try:
    from langchain_community.chains import ConversationalRetrievalChain
    print("FOUND in langchain_community.chains")
except ImportError:
    print("NOT FOUND in langchain_community.chains")

import langchain
print(f"Langchain file: {langchain.__file__}")

import os
langchain_dir = os.path.dirname(langchain.__file__)
print(f"Langchain contents: {os.listdir(langchain_dir)}")
