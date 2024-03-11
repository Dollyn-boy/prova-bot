from docx import Document

doc = Document()

# Listar os estilos disponíveis
print("Estilos Disponíveis:")
for style in doc.styles:
    print(style)