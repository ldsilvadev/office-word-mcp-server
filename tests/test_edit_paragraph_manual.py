import asyncio
import os
import sys

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from word_document_server.tools.content_tools import add_paragraph, edit_paragraph_text
from word_document_server.tools.document_tools import create_document, get_document_text

async def test_edit_paragraph():
    filename = "test_edit_paragraph.docx"
    
    # 1. Create document
    print(f"Creating {filename}...")
    await create_document(filename)
    
    # 2. Add paragraphs
    print("Adding paragraphs...")
    await add_paragraph(filename, "Paragraph 0")
    await add_paragraph(filename, "Paragraph 1")
    await add_paragraph(filename, "Paragraph 2")
    
    # 3. Verify initial state
    text = await get_document_text(filename)
    print(f"Initial text:\n{text}")
    assert "Paragraph 1" in text
    
    # 4. Edit paragraph 1
    print("Editing paragraph 1...")
    result = await edit_paragraph_text(filename, 1, "EDITED Paragraph 1")
    print(f"Result: {result}")
    
    # 5. Verify change
    text = await get_document_text(filename)
    print(f"Final text:\n{text}")
    
    if "EDITED Paragraph 1" in text and "Paragraph 1" not in text:
        print("SUCCESS: Paragraph edited successfully.")
    else:
        print("FAILURE: Paragraph not edited correctly.")

    # Cleanup
    if os.path.exists(filename):
        os.remove(filename)

if __name__ == "__main__":
    asyncio.run(test_edit_paragraph())
