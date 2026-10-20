# Leitor de NFs

Pequeno app que lê uma pasta de notas fiscais (PDF ou XML) e devolve uma planilha Excel com paciente, número da nota, valor bruto, valor líquido e mês de referência. Substituiu horas de copiar e colar.

> Código privado por enquanto; aqui fica o case.

## Como funciona
- Aceita arquivos soltos ou pastas inteiras, pela janela ou pela linha de comando
- **XML:** lê as tags do padrão nacional e de prefeituras (com nomes diferentes para o mesmo campo)
- **PDF:** extrai o texto e acha os valores com expressões regulares
- Se existem PDF e XML da mesma nota, prefere o XML (mais confiável)
- Avisa quando algo precisa de conferência (PDF escaneado, valor líquido não encontrado)
- Gera a planilha já formatada

## Stack
Python · PyMuPDF · XML (ElementTree) · Regex · openpyxl · Tkinter

## O que aprendi
- Extrair dados de documentos que não seguem um padrão só
- Usar regex para datas e valores em formato brasileiro
- Desenhar o programa para avisar quando não tem certeza, em vez de errar calado
