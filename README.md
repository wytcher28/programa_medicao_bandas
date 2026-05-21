# programa_medicao_bandas
Esse programa calculo a partir de um arquivo de logs, a linha de banda a ser utilziada para os fw PANW 
Programa medidas de calculo de banda 

Passo a passo 


1.) Dentro da pasta extraída, mantenha estes arquivos:
medidas_banda.py
requirements.txt
compilar_windows.bat
dados_log_bruto.xlsx


2.) Sempre que tiver uma nova planilha, ela deve se chamar exatamente:
dados_log_bruto.xlsx

3.) A planilha precisa ter a aba:
Dados

4.) Dentro da aba Dados, precisam existir estas colunas:
Bytes
Sessions
Application

5.) Para testar sem gerar executável
Abra o CMD dentro da pasta do projeto e rode:
pip install -r requirements.txt
python medidas_banda.py


6.) Será gerado o arquivo:
dados_log_bruto_resultado.xlsx
