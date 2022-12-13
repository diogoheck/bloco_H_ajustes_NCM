from openpyxl import Workbook
from openpyxl import load_workbook
from planilha_criacao import criacao_pastas_planilhas as criar
from leitura_arquivos import ler_sped_fiscal as leitura
from salvar_arquivos_pasta_trabalho import salvar_arquivos_0200 as salvar
from dicionarios import dic_0200 as criar_dict
from leitura_planilha_ncm import leitura_ncm as ler
import os

PASTA_DE_TRABALHO = 'Bloco_H.xlsx'
PLANILHA_0200 = '0200'
PLANILHA_H010 = 'H010'
PLANILHA_H020 = 'H020'
PLANLHA_SPED = 'SPED'
PLANILHA_NF_CREDITO = 'NF_CREDITO'
ARQUIVO_SPED = 'sped.txt'
ARQUIVO_NCMS = 'banco_dados_ncm.xlsx'


if __name__ == '__main__':

    # Criacao Dicionario 0200
    print('\niniciando...........\n')
    dic_itens = criar_dict.dict_arquivo_itens(ARQUIVO_SPED)
    dic_c170 = criar_dict.dict_arquivo_c170(ARQUIVO_SPED)
    dic_1923 = criar_dict.dict_arquivo_1923(ARQUIVO_SPED)
    dic_h010 = criar_dict.dict_arquivo_h010(ARQUIVO_SPED)
    dic_k200 = criar_dict.dict_arquivo_k200(ARQUIVO_SPED)
    # print(dic_k200)
    # Criar Planilhas
    print('criando pastas de trabalho e planilhas .....\n')
    criar.criacao_pasta_trabalho(PASTA_DE_TRABALHO)
    criar.criacao_planilhas(PLANILHA_0200, PLANILHA_H010,
                            PLANILHA_H020, PLANILHA_NF_CREDITO, PLANLHA_SPED)
    # leitura sped
    print('lendo Sped Fiscal........\n')
    lista_sped_fiscal = leitura.fazer_leitura_arquivo_sped(ARQUIVO_SPED)

    print('lendo NCMs validas......\n')
    # leitura do banco de dados de NCMs
    lista_ncms_validas = ler.leitura_ncm_validas(ARQUIVO_NCMS)
    # print(lista_ncms_validas)
    # leitura e geracao
    print('lendo 0200.........\n')
    dic_0200 = leitura.ler_arquivo_0200(
        lista_sped_fiscal, lista_ncms_validas, dic_c170, dic_1923, dic_h010, dic_k200)
    # print(dic_0200)
    # for k, v in dic_0200.items():
        # print(k)

    print('lendo h010 e h020............\n')
    lista_h010_h020 = leitura.ler_arquivo_h010_h020(
        lista_sped_fiscal, dic_itens, dic_0200)

    # salvar planilhas
    print('criando e salvando planilhas 0200.....\n')
    salvar.salvar_0200(PASTA_DE_TRABALHO, dic_0200, PLANILHA_0200)
    print('criando e salvando planilhas h010.....\n')
    salvar.salvar_h010(PASTA_DE_TRABALHO, lista_h010_h020, PLANILHA_H010)
    print('criando e salvando planilhas h020.....\n')
    salvar.salvar_h020(PASTA_DE_TRABALHO, lista_h010_h020, PLANILHA_H020)
    print('criando e salvando planilhas NF Credito.....\n')
    salvar.salvar_NF_CREDITO(PASTA_DE_TRABALHO, PLANILHA_NF_CREDITO)
    print('criando e salvando SPED.....\n')
    salvar.salvar_sped_0200(PASTA_DE_TRABALHO, PLANLHA_SPED, PLANILHA_0200)
    salvar.salvar_sped_h010_h020(PASTA_DE_TRABALHO, PLANLHA_SPED,
                                 PLANILHA_H010, PLANILHA_H020)

    print('Gerado com sucesso!')

    os.system('Pause')
