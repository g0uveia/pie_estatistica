###################################################
##    PIEDADE - FORMATA PLANILHAS ESTATISTICA    ##
##-----------------------------------------------##
##    4.PY - AUTOR: AUGUSTO GOUVEIA			   	 ##
##    Desc: Substitui caracteres	             ##
##												 ##
###################################################

import pandas as pd
import sys

filename = sys.argv[1]

df = pd.read_excel(filename + '.xlsx', header=None)

df = df.replace(to_replace='*', value='@')
df = df.replace(to_replace='%', value='')

print(df.head())




df.to_excel(filename + '.xlsx', index=False, header=False)