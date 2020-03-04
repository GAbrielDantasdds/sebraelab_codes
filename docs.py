from openpyxl import load_workbook



def positions(ws):
    try:
        dic = {}
        for valor in ws.values:
            if valor[1] != None and valor[1] != 'Email':
                dic[valor[0].title()] = {'email': valor[1], 'Telefone': valor[2], 'CPF': valor[3]}
        return dic
    except:
        return None


class documento():
    def __init__(self, caminho):
        self.arquivo = caminho
        self.data = {}

    def get_values(self):
        wb = load_workbook(self.arquivo)
        ws = wb['Lista de participantes']
        self.data = positions(ws)
        return self.data


arquivo = documento('lista semtepi).xlsx')
dados = arquivo.get_values()

with open('arquivo_temporario.csv', 'a') as arquivo:
    for chave, valor in dados.items():
        arquivo.write(f"{chave};{valor['email']}; {valor['Telefone']}\n")
