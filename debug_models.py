import anthropic
import os

api_key = os.environ.get("ANTHROPIC_API_KEY")
if not api_key:
    raise SystemExit("Set the ANTHROPIC_API_KEY environment variable before running this script.")

client = anthropic.Anthropic(api_key=api_key)

print("Searching for available models...")
try:
    for m in client.models.list():
        print(f"- {m.id}")
except Exception as e:
    print(f"Error: {e}")