import pypandoc
import os

def convert_md_to_docx():
    input_file = 'project_documentation.md'
    output_file = 'Multi-PDF_Chat_Documentation.docx'
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found!")
        return
    
    try:
        # Download pandoc if not available
        print("Checking for pandoc...")
        try:
            pypandoc.get_pandoc_version()
        except OSError:
            print("Pandoc not found. Downloading pandoc...")
            pypandoc.download_pandoc()
            print("Pandoc downloaded successfully!")

        # Convert markdown to docx
        print("Converting markdown to DOCX...")
        pypandoc.convert_file(
            input_file,
            'docx',
            outputfile=output_file,
            format='markdown'
        )
        print(f"Successfully converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Error during conversion: {str(e)}")

if __name__ == "__main__":
    convert_md_to_docx() 