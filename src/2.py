####################################################
##    PIEDADE - FORMATA PLANILHAS ESTATISTICA     ##
##------------------------------------------------##
##    2.PY - AUTOR: AUGUSTO GOUVEIA			   	  ##
##    INFO: Remove colunas vazias		          ##
##												  ##
####################################################

import sys
import pandas as pd

filename = sys.argv[1]

### CARREGA ARQUIVO
df = pd.read_excel(filename + '.xlsx', header=None)

df = df.replace(to_replace='Respostas', value=None)

### REMOVE COLUNAS VAZIAS
df = df.dropna(how='all', axis=1)

### SALVA ARQUIVO
df.to_excel(filename + '.xlsx', index=False, header=False)
