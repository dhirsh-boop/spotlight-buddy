import google.generativeai as genai
import os

# Paste your API Key here just for this test
api_key = "AIzaSyB3kkk1T9e8vLeFhTseMRaxrxsxIrTlILE"

genai.configure(api_key=api_key)

print("Searching for available models...")
try:
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(f"- {m.name}")
except Exception as e:
    print(f"Error: {e}")