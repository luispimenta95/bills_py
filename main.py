import pandas as pd

class Main:
    def __init__(self, caminho_arquivo, inicial):
        self.caminho_arquivo = caminho_arquivo
        self.inicial = inicial
        self.maior = 0
        self.menor = float('inf')
        self.parcela_maior = None
        self.parcela_menor = None
        self.valores_pagos = []
        self.data_ultima_em_dia = None  # Variável para a data da última parcela "EM DIA"

    def converterValores(self, valor):
        return f'R$ {valor:,.2f}'.replace('.', ',').replace(',', '.', 1)

    def processar_dados(self):
        # Leitura do arquivo Excel
        df = pd.read_excel(self.caminho_arquivo)
        df = df.iloc[:-1]  # Exclui a última linha

        # Inicializa uma variável para a última data de vencimento "EM DIA"
        ultima_data_em_dia = None

        # Percorrer as linhas e detalhar as colunas
        for index, row in df.iterrows():
            valor_pago = row.get("Valor pago")
            desconto = row.get("Desconto")
            parcela = row.get("Parcela")  # Obtém o número da parcela
            status = row.get("Situação")  # Obtém o status da parcela
            vencimento = row.get("Vencimento")  # Obtém a data de vencimento da parcela

            # Converte valor para float se possível
            try:
                if valor_pago:
                    valor_pago = float(valor_pago)
                if desconto:
                    desconto = float(desconto)
            except ValueError:
                continue  # Ignora valores que não podem ser convertidos

            # Atualiza menor valor e sua parcela
            if valor_pago < self.menor:
                self.menor = valor_pago
                self.parcela_menor = parcela

            # Atualiza maior valor e sua parcela
            if valor_pago > self.maior and status == 'PAGO':
                self.maior = valor_pago
                self.parcela_maior = parcela

            # Adiciona o valor pago à lista se o status for 'PAGO'
            if status == 'PAGO':
                self.valores_pagos.append(valor_pago)

            # Atualiza a data de vencimento da última parcela 'EM DIA'
            if status == 'EM DIA':
                ultima_data_em_dia = pd.to_datetime(vencimento, format='%d/%m/%Y')

        # Define a data da última parcela 'EM DIA' após o loop
        if ultima_data_em_dia is not None:
            self.data_ultima_em_dia = ultima_data_em_dia

    def calcular_resultados(self):
        # Calcula a média dos valores pagos
        if self.valores_pagos:
            qtd_pg = len(self.valores_pagos)
            media_pagos = sum(self.valores_pagos) / qtd_pg
        else:
            qtd_pg = 0
            media_pagos = 0

        pg_desconto = (self.inicial * qtd_pg) - sum(self.valores_pagos)
        parcelas_economizadas = int(pg_desconto / self.inicial)

        # Formata os valores
        self.maior = self.converterValores(self.maior)
        self.menor = self.converterValores(self.menor)
        media_pagos = self.converterValores(media_pagos)
        pg_desconto = self.converterValores(pg_desconto)

        # Retorna os resultados formatados
        return {
            'maior': self.maior,
            'parcela_maior': self.parcela_maior,
            'menor': self.menor,
            'parcela_menor': self.parcela_menor,
            'qtd_pg': qtd_pg,
            'media_pagos': media_pagos,
            'pg_desconto': pg_desconto,
            'parcelas_economizadas': parcelas_economizadas,
            'data_ultima_em_dia': self.data_ultima_em_dia.strftime('%d/%m/%Y') if self.data_ultima_em_dia else 'N/A'
        }

    def exibir_resultados(self):
        resultados = self.calcular_resultados()
        largura = 85
        borda = "#" * largura
        numParcelas = 50

        print(borda)
        print(f'Valor inicial da parcela: {self.converterValores(self.inicial)}')
        print(f'Maior valor: {resultados["maior"]} na parcela {resultados["parcela_maior"]}')
        print(f'Menor valor: {resultados["menor"]} na parcela {resultados["parcela_menor"]}')
        print(f'Quantidade de parcelas pagas: {resultados["qtd_pg"]}')
        print(f'Restam {numParcelas - resultados["qtd_pg"]} parcelas para serem pagas')
        print(f'Data provável para quitar a divida: {resultados["data_ultima_em_dia"]}')

        print(f'Média do valor pago nas parcelas: {resultados["media_pagos"]}')
        print(f'Total de desconto: {resultados["pg_desconto"]}')
        print(f'Quantidade de parcelas pagas com o desconto de {resultados["pg_desconto"]}: {resultados["parcelas_economizadas"]}')
        print(borda)

if __name__ == "__main__":
    # Caminho para o seu arquivo Excel e valor inicial
    caminho_arquivo = 'fin.xlsx'
    inicial = 2256.61

    # Crie uma instância da classe Main e execute o processo
    main = Main(caminho_arquivo, inicial)
    main.processar_dados()
    main.exibir_resultados()
