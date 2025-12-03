"""
Script para testar e investigar tags no documento Word.
Verifica quais tags estão presentes na primeira linha de dados das tabelas.
"""
from docx import Document
import re
import sys

def listar_tags_no_documento(doc_path):
    """Lista todas as tags encontradas no documento."""
    doc = Document(doc_path)
    
    print("=" * 80)
    print(f"INVESTIGAÇÃO DE TAGS - {doc_path}")
    print("=" * 80)
    print()
    
    # Procura em todas as tabelas
    for table_idx, table in enumerate(doc.tables):
        print(f"\n{'='*80}")
        print(f"TABELA {table_idx}")
        print(f"{'='*80}")
        
        # Verifica cada linha da tabela
        for row_idx, row in enumerate(table.rows):
            texto_linha_completo = ' '.join([cell.text.strip() for cell in row.cells])
            
            # Procura por tags com padrão {TAG}
            tags_encontradas = re.findall(r'\{[A-Z_0-9]+\}', texto_linha_completo)
            
            if tags_encontradas:
                print(f"\n  Linha {row_idx}:")
                print(f"    Texto completo: {texto_linha_completo[:100]}...")
                print(f"    Tags encontradas: {tags_encontradas}")
                
                # Mostra conteúdo de cada célula
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text and ('{' in cell_text or '}' in cell_text):
                        print(f"      Célula {cell_idx}: {cell_text}")
    
    print("\n" + "=" * 80)
    print("FIM DA INVESTIGAÇÃO")
    print("=" * 80)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python testar_tags.py <caminho_do_documento.docx>")
        print("\nExemplo:")
        print("  python testar_tags.py \"Modelo PT-CURSOR.docx\"")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    
    try:
        listar_tags_no_documento(doc_path)
    except FileNotFoundError:
        print(f"ERRO: Arquivo não encontrado: {doc_path}")
        sys.exit(1)
    except Exception as e:
        print(f"ERRO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

