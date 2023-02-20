# Projeto - Relatório de fechamento de mercado por e-mail
# Construir um e-mail que chegue na caixa de entrada todos os dias com informações de fechamentos do Ibovespa e Dolar


# Importando os módulos necessários

import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as win32

# Pegar dados no Yahoo Finance

codigos_de_negociacao = ['^BVSP', 'BRL=X']

hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)

dados_mercado = yf.download(codigos_de_negociacao, um_ano_atras, hoje)

print(dados_mercado)

# Manipulando os dados - seleção e exclusão de dados

dados_fechamento = dados_mercado['Adj Close']

dados_fechamento.columns = ['dolar', 'ibovespa']

dados_fechamento = dados_fechamento.dropna()

print(dados_fechamento.head(50))

# Manipulando os dados - Criando tabelas com outros timeframes

dados_anuais = dados_fechamento.resample('Y').last() #sum = soma. last é o fechamento

dados_mensais = dados_fechamento.resample('M').last()

print(dados_anuais)

# Calcular fechamento do dia, retorno no ano e retorno no mês dos ativos

retorno_anual = dados_anuais.pct_change().dropna()
retorno_mensal = dados_mensais.pct_change().dropna()
retorno_diario = dados_fechamento.pct_change().dropna()

print(retorno_anual)
print(retorno_mensal)
print(retorno_diario)

# Localizar o fechamento do dia anterior, retorno no mês e retorno no ano e retorno no mês.
# loc -> referenciar elementos a partir do nome ou iloc -> selecionar elementos como uma matriz

# retorno_jan_26_2022 = retorno_diario.loc['2022-10-26', 'dolar']
# retorno_jan_26_2022_iloc = retorno_diario.iloc[1, 0]


retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual.iloc[-1, 0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]

retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual.iloc[-1, 0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]

print(retorno_diario)

#em decimais

retorno_diario_dolar = round((retorno_diario_dolar * 100), 2)
retorno_diario_ibov = round((retorno_diario_ibov * 100), 2)

retorno_mensal_dolar = round((retorno_mensal_dolar * 100), 2)
retorno_mensal_ibov = round((retorno_mensal_ibov * 100), 2)

retorno_anual_dolar = round((retorno_anual_dolar * 100), 2)
retorno_anual_ibov = round((retorno_anual_ibov * 100), 2)

print(retorno_anual_ibov)


# Fazer os gráficos da performance do último dos ativos

plt.style.use('cyberpunk')

dados_fechamento.plot(y = 'ibovespa', use_index = True, legend = False)

plt.title('Ibovespa')

plt.show()

#dolar

plt.style.use('cyberpunk')

dados_fechamento.plot(y = 'dolar', use_index = True, legend = False)

plt.title('dolar')

plt.show()

# Enviar e-mail

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = "lucaskevin455@gmail.com"
email.Subject = "Relatório Diário"
email.Body = f'''Prezado diretor, segue o relatório diário:

Bolsa:

No ano o Ibovespa está tendo uma rentabilidade de {retorno_anual_ibov}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_ibov}%.

No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

Dólar:

No ano o Dólar está tendo uma rentabilidade de {retorno_anual_dolar}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_dolar}%.

No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.


Abs,

O melhor estagiário do mundo

'''

anexo_ibovespa = r'C:\Users\luqui\OneDrive\Documentos\Estudos\bootcamp-py\Aula1\ibovespa.png'
anexo_dolar = r'C:\Users\luqui\OneDrive\Documentos\Estudos\bootcamp-py\Aula1\dolar.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)

email.Send()