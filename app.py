import fitz  # PyMuPDF
import sys
import os
from xlsreader import read_excel_data, get_tooltip_content

def process_pdf_tooltips(input_pdf_path, output_pdf_path=None, excel_data=None):
    """
    Read tooltip content from PDF and replace with Excel data.
    
    Args:
        input_pdf_path (str): Path to the input PDF file
        output_pdf_path (str): Path for the output PDF file (optional)
        excel_data (dict): Dictionary containing Excel data by row
    """
    
    # Check if input file exists
    if not os.path.exists(input_pdf_path):
        print(f"Error: File '{input_pdf_path}' not found!")
        return False
    
    # Set default output path if not provided
    if output_pdf_path is None:
        base_name = os.path.splitext(input_pdf_path)[0]
        output_pdf_path = f"{base_name}_modified.pdf"
    
    try:
        # Open the PDF document
        doc = fitz.open(input_pdf_path)
        print(f"Successfully opened: {input_pdf_path}")
        print(f"Total pages: {len(doc)}")
        print("=" * 60)
        
        tooltip_count = 0
        
        # Process each page
        for page_num in range(len(doc)):
            page = doc[page_num]
            annotations = page.annots()
            
            if annotations:
                print(f"\nPage {page_num + 1}:")
                print("-" * 40)
                
                # Process each annotation
                for annot_index, annot in enumerate(annotations):
                    # Check if it's a tooltip annotation
                    annot_type = annot.type[1]
                    
                    if annot_type in ["Text", "Note", "FreeText", "Popup", "Highlight"]:
                        # Read current tooltip content
                        current_content = annot.info.get("content", "")
                        
                        # Get new content from Excel data
                        new_content = get_tooltip_content(excel_data, current_content)
                        
                        # Print current and new content
                        print(f"  Annotation {annot_index + 1} ({annot_type}):")
                        print(f"    Current content: '{current_content}'")
                        
                        if new_content:
                            # Update tooltip content
                            annot.set_info(content=new_content)
                            annot.update()
                            print(f"    New content: '{new_content}'")
                        else:
                            print("    No matching Excel data found - keeping original content")
                        
                        print()
                        tooltip_count += 1
        
        print("=" * 60)
        print(f"Total tooltips processed: {tooltip_count}")
        
        if tooltip_count > 0:
            # Save the modified PDF
            doc.save(output_pdf_path)
            print(f"Modified PDF saved as: {output_pdf_path}")
        else:
            print("No tooltips found in the PDF.")
        
        # Close the document
        doc.close()
        return True
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return False

def main():
    """
    Main function to run the tooltip processor
    """
    # File paths
    excel_file = "Perla_11_1_vorlage-steps.xlsx"  # Excel file path
    input_pdf = "template.pdf"  # Change this to your PDF file path
    output_pdf = "output_modified.pdf"  # Change this to your desired output path
    
    # Alternative: Get file paths from command line arguments
    if len(sys.argv) > 1:
        input_pdf = sys.argv[1]
    if len(sys.argv) > 2:
        output_pdf = sys.argv[2]
    if len(sys.argv) > 3:
        excel_file = sys.argv[3]
    
    print("PDF Tooltip Reader and Editor")
    print("=" * 60)
    print(f"Input PDF: {input_pdf}")
    print(f"Output PDF: {output_pdf}")
    print(f"Excel file: {excel_file}")
    print()
    
    try:
        # Read Excel data
        print("Reading Excel data...")
        excel_data = read_excel_data(excel_file)
        
        # Process the PDF with Excel data
        success = process_pdf_tooltips(input_pdf, output_pdf, excel_data)
        
        if success:
            print("\nProcess completed successfully!")
        else:
            print("\nProcess failed!")
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()