import pandas as pd

class Proc:
    def __init__(self):
        caminho_arquivo = 'fin.xlsx'
        inicial = 2256.61
        self.caminho_arquivo = caminho_arquivo
        self.inicial = inicial
        self.maior = 0
        self.menor = float('inf')
        self.parcela_maior = None
        self.parcela_menor = None
        self.valores_pagos = []
        self.data_ultima_em_dia = None
        self.data_em_dia = None  # Inicializa a variável para a primeira data "EM DIA"

    def converterValores(self, valor):
        return f'R$ {valor:,.2f}'.replace('.', ',').replace(',', '.', 1)

    def processar_dados(self):
        df = pd.read_excel(self.caminho_arquivo)
        df = df.iloc[:-1]

        ultima_data_em_dia = None

        for index, row in df.iterrows():
            valor_pago = row.get("Valor pago")
            desconto = row.get("Desconto")
            parcela = row.get("Parcela")
            status = row.get("Situação")
            vencimento = row.get("Vencimento")

            try:
                if valor_pago:
                    valor_pago = float(valor_pago)
                if desconto:
                    desconto = float(desconto)
            except ValueError:
                continue

            if valor_pago < self.menor:
                self.menor = valor_pago
                self.parcela_menor = parcela

            if valor_pago > self.maior and status == 'PAGO':
                self.maior = valor_pago
                self.parcela_maior = parcela

            if status == 'PAGO':
                self.valores_pagos.append(valor_pago)

            if status == 'EM DIA':
                ultima_data_em_dia = pd.to_datetime(vencimento, format='%d/%m/%Y')
                if self.data_em_dia is None:
                    self.data_em_dia = ultima_data_em_dia  # A primeira data 'EM DIA'

        if ultima_data_em_dia is not None:
            self.data_ultima_em_dia = ultima_data_em_dia

    def calcular_resultados(self):
        if self.valores_pagos:
            qtd_pg = len(self.valores_pagos)
            media_pagos = sum(self.valores_pagos) / qtd_pg
        else:
            qtd_pg = 0
            media_pagos = 0

        pg_desconto = (self.inicial * qtd_pg) - sum(self.valores_pagos)
        parcelas_economizadas = int(pg_desconto / self.inicial)

        self.maior = self.converterValores(self.maior)
        self.menor = self.converterValores(self.menor)
        media_pagos = self.converterValores(media_pagos)
        pg_desconto = self.converterValores(pg_desconto)

        return {
            'maior': self.maior,
            'parcela_maior': self.parcela_maior,
            'menor': self.menor,
            'parcela_menor': self.parcela_menor,
            'qtd_pg': qtd_pg,
            'media_pagos': media_pagos,
            'pg_desconto': pg_desconto,
            'parcelas_economizadas': parcelas_economizadas,
            'data_ultima_em_dia': self.data_ultima_em_dia.strftime('%d/%m/%Y') if self.data_ultima_em_dia  else 'Todas parcelas foram pagas',
            'proximo_vencimento': self.data_em_dia.strftime("%d/%m/%Y") if self.data_em_dia else 'Todas parcelas foram pagas',
        }

    def exibir_resultados(self):
        resultados = self.calcular_resultados()
        largura = 85
        borda = "#" * largura
        numParcelas = 50

        #print(borda)
        print(f'Valor inicial da parcela: {self.converterValores(self.inicial)}')
        print(f'Maior valor: {resultados["maior"]} na parcela {resultados["parcela_maior"]}')
        print(f'Menor valor: {resultados["menor"]} na parcela {resultados["parcela_menor"]}')
        print(f'Quantidade de parcelas pagas: {resultados["qtd_pg"]}')
        print(f'Restam {numParcelas - resultados["qtd_pg"]} parcelas para serem pagas')
        print(f'Data da próxima parcela: {resultados["proximo_vencimento"]}')

        print(f'Data provável para quitar a dívida: {resultados["data_ultima_em_dia"]}')

        print(f'Média do valor pago nas parcelas: {resultados["media_pagos"]}')
        print(f'Total de desconto: {resultados["pg_desconto"]}')
        print(f'Quantidade de parcelas pagas com o desconto de {resultados["pg_desconto"]}: {resultados["parcelas_economizadas"]}')
        #print(borda)
