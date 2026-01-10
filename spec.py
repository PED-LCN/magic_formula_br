import fundamentus
import pandas as pd

def inspecionar():
    print("1. Baixando dados brutos...")
    df = fundamentus.get_resultado()

    print("\n--- A. QUAIS SÃO AS COLUNAS? ---")
    print(list(df.columns))

    print("\n--- B. COMO O PYTHON ESTÁ LENDO OS DADOS? (Tipos) ---")
    # Se aparecer 'object', o Python acha que é Texto. 
    # Se aparecer 'float64', ele sabe que é Número.
    print(df.dtypes)

    print("\n--- C. AMOSTRA (PRIMEIRA LINHA) ---")
    # Mostra a primeira empresa para você ver a formatação (vírgula, ponto, %, R$)
    print(df.iloc[0])

    # Gera um arquivo bruto para você abrir no Excel e olhar com calma
    print("\n--- D. SALVANDO ARQUIVO DE DEBUG ---")
    df.to_csv('debug_dados_brutos.csv', sep=';', encoding='utf-8-sig')
    print("Arquivo 'debug_dados_brutos.csv' gerado na pasta!")

if __name__ == "__main__":
    inspecionar()