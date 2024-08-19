import win32com.client as win32
import os

def create_new_page_in_word(file_path=None, save_as=None, visible=True):
    try:
        # Initialize Word application
        word = win32.Dispatch('Word.Application')
        word.Visible = visible  # Set visibility

        # Open an existing document or create a new one
        if file_path and os.path.exists(file_path):
            doc = word.Documents.Open(file_path)
        else:
            doc = word.Documents.Add()
        
        # Move the cursor to the end of the document
        doc.Range().Collapse(Direction=0)  # 0 for wdCollapseEnd

        # Insert a page break (new page)
        doc.Range().InsertBreak(Type=7)  # 7 for wdPageBreak

        # Determine the save path
        if save_as:
            save_path = save_as
        else:
            save_path = file_path if file_path else "NewDocument.docx"
        
        # Save the document
        doc.SaveAs(save_path)

        print(f"New page added and document saved as: {save_path}")

    except Exception as e:
        print(f"An error occurred: {e}")
    # Removed the finally block to keep the Word application open
    # The document and Word application will stay open

# Example usage
create_new_page_in_word()
