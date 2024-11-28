from pptx import Presentation

# Load the presentation template

prs = Presentation('C:/Users/raphael.dias/OneDrive - DINACO IMPORTACAO COMERCIO S A/Documentos Compartilhados - Intranet Dinaco/Inteligência Comercial/Documentos/Automação/Suporte/Template_Gerencia.pptx')

# Choose the slide layout index to use
layout_index = 3

# Get the layout
slide_layout = prs.slide_layouts[layout_index]

# Print Placeholder details for the chosen layout
print(f"Details for layout {layout_index}: {slide_layout.name}")
for placeholder in slide_layout.placeholders:
    print(f"Placeholder index: {placeholder.placeholder_format.idx}, Type: {placeholder.placeholder_format.type}, Name: '{placeholder.name}'")