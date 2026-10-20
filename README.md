# Leitor de NFs

Pequeno app que lê notas fiscais (DANFe/NFS-e) em PDF ou XML e devolve uma planilha Excel com paciente (tomador/beneficiário), número da nota, valor bruto, valor líquido e mês de referência. Substituiu horas de copiar e colar.

## Como funciona
- Aceita arquivos soltos ou pastas inteiras, pela janela ou pela linha de comando
- **XML:** lê as tags do padrão nacional e de prefeituras (que usam nomes diferentes para o mesmo campo)
- **PDF:** extrai o texto e acha os valores com expressões regulares
- Se existem PDF e XML da mesma nota, prefere o XML (mais confiável)
- Avisa quando algo precisa de conferência, como PDF escaneado ou valor líquido não encontrado
- Gera a planilha já formatada

## Como rodar
Precisa de **Python 3.10+**.

```bash
pip install -r requirements.txt
python leitor_nf.py
```

No Windows, também dá para abrir com dois cliques em `Abrir_Leitor_NF.bat`.

Pela linha de comando, passe arquivos ou pastas:

```bash
python leitor_nf.py nota1.pdf nota2.xml pasta_com_notas
```

## Stack
Python · PyMuPDF · XML (ElementTree) · Regex · openpyxl · Tkinter

## O que aprendi
- Extrair dados de documentos que não seguem um padrão só
- Usar regex para datas e valores no formato brasileiro
- Desenhar o programa para avisar quando não tem certeza, em vez de errar calado

## Licença
[MIT](LICENSE)
