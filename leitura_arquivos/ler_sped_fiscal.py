


ARQUIVO_SPED = 'sped.txt'


def isnumber(valor):
    try:
        float(valor)
    except ValueError:
        return False
    return True


def converte_valor(texto):
    if isnumber(texto):
        valor = texto
    elif '.' in texto and ',' in texto:
        valor = texto.replace('.', '').replace(',', '.')
    else:
        valor = texto.replace(',', '.')
    try:
        return round(float(valor), 2)
    except:

        return valor


def verifica_ncm_valida(s, lista_ncms_validas):

    
    if s[8][0:4] in lista_ncms_validas or \
        s[8][0:5] in lista_ncms_validas or \
        s[8][0:6] in lista_ncms_validas or \
        s[8][0:7] in lista_ncms_validas or \
            s[8][0:8] in lista_ncms_validas:
        if s[8] == '1002000':
            print(s[8])
        s[14] = 'S'


def tem_cadastro_c100(dic_170, cod_item_0200, s):
    if dic_170.get(cod_item_0200):
        s[15] = 'S'


def tem_cadastro_1923(dic_1923, cod_item_0200, s):
    if dic_1923.get(cod_item_0200):
        s[16] = 'S'


def tem_cadastro_h010(dic_h010, cod_item_0200, s):
    if dic_h010.get(cod_item_0200):
        s[17] = 'S'


def tem_cadastro_k200(dic_k200, cod_item_0200, s):
    if dic_k200.get(cod_item_0200):
        s[18] = 'S'


def ler_arquivo_0200(sped_fiscal, lista_ncms_validas, dic_c170, dic_1923, dic_h010, dic_k200):
    lista_0200 = []
    dic_0200 = {}
    i = 0
    for s in sped_fiscal:
        if s[1] == '0200' or s[1] == '0220':

            if s[1] == '0200':
                # print(s)
                # print(s)
                s[12] = converte_valor(s[12]) if s[12] else 0.0
                s[12] = 17 if s[12] == 0 else s[12]
                # print(s[12], type(s[12]))
                # 14 -> ncm, 15 -> c100
                s[14] = 'N'
                s.append('N')
                s.append('N')
                s.append('N')
                s.append('N')
                verifica_ncm_valida(s, lista_ncms_validas)
                tem_cadastro_c100(dic_c170, s[2], s)
                tem_cadastro_1923(dic_1923, s[2], s)
                tem_cadastro_h010(dic_h010, s[2], s)
                tem_cadastro_k200(dic_k200, s[2], s)

            # print(s[8])
            # print(s)
            if s[1] == '0220':
                dic_0200[s[2] + str(i)] = s
                i += 1
            else:
                dic_0200[s[2]] = s

            # lista_0200.append(s)
    # print(dic_0200)
    return dic_0200


class BlocoH:
    def __init__(self) -> None:
        self.lista_blocos = []

    def add_bloco(self, bloco):
        self.lista_blocos.append(bloco)


class H10:
    def __init__(self) -> None:
        self.lista_h010 = []
        self.lista_h020 = []

    def add_lista_h010(self, campos):
        self.lista_h010 = campos

        self.lista_h010[4] = converte_valor(self.lista_h010[4])
        self.lista_h010[5] = converte_valor(self.lista_h010[5])
        self.lista_h010[6] = converte_valor(self.lista_h010[6])
        self.lista_h010[11] = converte_valor(self.lista_h010[11])
        self.lista_h020 = ['', 'H020', '060',
                           self.lista_h010[5], self.lista_h010[5], '', '']

    def add_lista_h020(self, campos):
        self.lista_h020 = campos
        self.lista_h020[3] = converte_valor(self.lista_h020[3])
        self.lista_h020[3] = self.lista_h010[5] if self.lista_h020[3] == 0 else self.lista_h020[3]
        self.lista_h020[4] = converte_valor(self.lista_h020[4])


def ler_arquivo_h010_h020(sped_fiscal, dic_itens, dic_0200):
    lista_geral = []

    total_blocos = BlocoH()
    bloco_h = H10()
    lista_itens_add_h020 = []
    contador = 0
    for s in sped_fiscal:

        if s[1] == 'H010' or s[1] == 'H020':

            if s[1] == 'H010':
                # print(dic_0200.get(s[2]))
                if dic_0200.get(s[2]) is not None:
                    if dic_0200.get(s[2])[14] == 'S':
                        bloco_h = H10()
                        bloco_h.add_lista_h010(s)
                        total_blocos.add_bloco(bloco_h)
                    ultima_ref = s

            if s[1] == 'H020':
                if ultima_ref:
                    # print(dic_0200.get(ultima_ref[2]))
                    if dic_0200.get(ultima_ref[2])[14] == 'S':
                        if ultima_ref[2] not in lista_itens_add_h020:
                            lista_itens_add_h020.append(ultima_ref[2])
                            bloco_h.add_lista_h020(s)

    for lista in total_blocos.lista_blocos:

        lista_geral.append([lista.lista_h010, lista.lista_h020])

    return lista_geral


def fazer_leitura_arquivo_sped(ARQUIVO_SPED):
    sped_fiscal = []
    with open(ARQUIVO_SPED, encoding='ansi') as arquivo:
        for registro in arquivo:
            campos = registro.strip().split('|')
            if campos[1] == '9999':
                break
            sped_fiscal.append(campos)
    return sped_fiscal


if __name__ == '__main__':
    sped_fiscal = fazer_leitura_arquivo_sped(ARQUIVO_SPED)
    lista_0200 = ler_arquivo_0200(sped_fiscal)
    print(lista_0200)
