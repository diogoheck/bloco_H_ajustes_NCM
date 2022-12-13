def dict_arquivo_itens(arquivo):
    with open(arquivo, encoding='ansi') as arquivo:
        nf = {}
        for registro in arquivo:
            produtos = registro.strip().split('|')
            if produtos[1] == '9999':
                break
            if produtos[1] == '0200' or produtos[1] == '0220':
                # print(produtos)
                nf[produtos[2]] = produtos

    return nf


def dict_arquivo_c170(arquivo):
    with open(arquivo, encoding='ansi') as arquivo:
        nf = {}
        for registro in arquivo:
            produtos = registro.strip().split('|')
            if produtos[1] == '9999':
                break
            if produtos[1] == 'C170':
                # print(produtos)
                nf[produtos[3]] = produtos
    return nf


def dict_arquivo_1923(arquivo):
    with open(arquivo, encoding='ansi') as arquivo:
        nf = {}
        for registro in arquivo:
            produtos = registro.strip().split('|')
            if produtos[1] == '9999':
                break
            if produtos[1] == '1923':
                # print(produtos)
                nf[produtos[8]] = produtos
    return nf


def dict_arquivo_h010(arquivo):
    with open(arquivo, encoding='ansi') as arquivo:
        nf = {}
        for registro in arquivo:
            produtos = registro.strip().split('|')
            if produtos[1] == '9999':
                break
            if produtos[1] == 'H010':
                # print(produtos)
                nf[produtos[2]] = produtos
    return nf


def dict_arquivo_k200(arquivo):
    with open(arquivo, encoding='ansi') as arquivo:
        nf = {}
        for registro in arquivo:
            produtos = registro.strip().split('|')
            if produtos[1] == '9999':
                break
            if produtos[1] == 'K200':
                # print(produtos)
                nf[produtos[3]] = produtos
    return nf


if __name__ == '__main__':
    dict_arquivo_itens('sped.txt')
