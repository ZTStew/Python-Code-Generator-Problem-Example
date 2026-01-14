from docx import Document

def main():
  output = []
  codes = ['asdf', 'fdsa', '1234', '4567']
  
  i = 0
  file_count = 1
  # Loops through provided list of `codes`
  while i < len(codes):
    # Marks location of Template File
    doc = Document('./templates/Template_1.docx')

    # inserts codes into template file
    for p in doc.paragraphs:
      if '{code1}' in p.text:
        p.text = p.text.replace('{code1}', codes[i])
        output.append(f'Saved File: `./to_print/custom_file_{i}.docx` with code: `{i}`')
      try:
        # Saves filled Template File
        doc.save(f'./to_print/custom_file_{file_count}.docx')
      except Exception as e:
        print(e)
  
    i += 1
    file_count += 1

  print(output)


if __name__ == "__main__":
  main()
