from docx import Document
import os

class DocxTableConverter:
    def __init__(self):
        self.markdown_tables = []

    def convert_table_to_markdown(self, table):
        """Convert a single table to markdown format"""
        markdown_rows = []
        
        # Get headers
        headers = []
        for cell in table.rows[0].cells:
            headers.append(cell.text.strip())
        markdown_rows.append('| ' + ' | '.join(headers) + ' |')
        
        # Add separator row
        separator = '|' + '|'.join(['---' for _ in headers]) + '|'
        markdown_rows.append(separator)
        
        # Get data rows
        for row in table.rows[1:]:
            cells = []
            for cell in row.cells:
                # Replace any newlines in cells with spaces
                cell_text = cell.text.strip().replace('\n', ' ')
                cells.append(cell_text)
            markdown_rows.append('| ' + ' | '.join(cells) + ' |')
        
        return '\n'.join(markdown_rows)

    def convert_file(self, input_file, output_file):
        """Convert all tables in a DOCX file to markdown format"""
        try:
            # Check if input file exists
            if not os.path.exists(input_file):
                raise FileNotFoundError(f"Input file not found: {input_file}")

            # Load the document
            print(f"Converting tables from: {input_file}")
            doc = Document(input_file)
            
            # Convert each table
            for i, table in enumerate(doc.tables, 1):
                print(f"Converting table {i}...")
                markdown_table = self.convert_table_to_markdown(table)
                self.markdown_tables.append(f"\n### Table {i}\n\n{markdown_table}\n")
            
            # Write to output file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("# Converted Tables\n\n")
                f.write('\n'.join(self.markdown_tables))
            
            print(f"\nConversion completed! Output saved to: {output_file}")
            print(f"Found and converted {len(self.markdown_tables)} tables")

        except Exception as e:
            print(f"Error during conversion: {str(e)}")

def main():
    try:
        converter = DocxTableConverter()
        
        # Example usage
        input_file = "input.docx"  # Replace with your DOCX file
        output_file = "output.md"  # Replace with desired output path
        
        converter.convert_file(input_file, output_file)
        
    except Exception as e:
        print(f"\nError: {str(e)}")

if __name__ == "__main__":
    main() 