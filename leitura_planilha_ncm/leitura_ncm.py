from openpyxl import load_workbook


def leitura_ncm_validas(ARQUIVO_NCMS):
    # carregar arquivo
    arquivo_excel = ARQUIVO_NCMS
    arquivo_excel = load_workbook(arquivo_excel)
    arquivo_excel = arquivo_excel.active

    lista_ncms = [str(linha_arquivo[0])
                  for linha_arquivo in arquivo_excel.values]
    return lista_ncms
