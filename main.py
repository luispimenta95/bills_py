from processaXls import  Proc as construct

if __name__ == "__main__":
    # Caminho para o seu arquivo Excel e valor inicial
    caminho_arquivo = 'fin.xlsx'
    inicial = 2256.61

    main = construct(caminho_arquivo, inicial)
    main.processar_dados()
    main.exibir_resultados()
