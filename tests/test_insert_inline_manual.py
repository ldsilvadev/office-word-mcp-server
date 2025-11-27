import asyncio
import os
import sys

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from word_document_server.tools.content_tools import add_paragraph, insert_text_inline
from word_document_server.tools.document_tools import create_document, get_document_text

async def test_insert_inline():
    filename = "test_insert_inline.docx"
    
    # 1. Create document
    print(f"Creating {filename}...")
    await create_document(filename)
    
    # 2. Add paragraph
    print("Adding paragraph...")
    await add_paragraph(filename, "This is an example sentence.")
    
    # 3. Verify initial state
    text = await get_document_text(filename)
    print(f"Initial text:\n{text}")
    
    # 4. Insert text AFTER "example"
    print("Inserting text AFTER 'example'...")
    result = await insert_text_inline(filename, "example", " (inserted after)")
    print(f"Result: {result}")
    
    # 5. Insert text BEFORE "sentence"
    print("Inserting text BEFORE 'sentence'...")
    result = await insert_text_inline(filename, "sentence", "modified ", position="before")
    print(f"Result: {result}")
    
    # 6. Verify changes
    text = await get_document_text(filename)
    print(f"Final text:\n{text}")
    
    expected_text = "This is an example (inserted after) modified sentence."
    
    if expected_text in text:
        print("SUCCESS: Text inserted inline correctly.")
    else:
        print(f"FAILURE: Text not inserted correctly.\nExpected: {expected_text}\nGot: {text}")

    # Cleanup
    if os.path.exists(filename):
        os.remove(filename)

if __name__ == "__main__":
    asyncio.run(test_insert_inline())
