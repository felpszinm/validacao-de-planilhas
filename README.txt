============================================================
PROGRAMA DE REESCRITA E VALIDAÇÃO DE DADOS EM PLANILHA EXCEL
Autor: Felipe Santos
============================================================

DESCRIÇÃO
---------
Script em Python para automatizar a leitura, atualização e validação de informações em planilhas Excel (.xlsx ou .xlsm), preservando fórmulas originais e garantindo consistência nos dados.

FUNCIONALIDADES
---------------
- Leitura de dados calculados no Excel (data_only=True).
- Atualização de valores mantendo fórmulas originais.
- Conversão automática de datas para o formato brasileiro (DD/MM/YYYY).
- Preenchimento automático da aba CONTROLE_RENTALS com:
  - Data
  - Placa
  - Status de manutenção (SIM/NÃO)
- Interrupção automática ao detectar linha vazia na lista de rotas.
- Validação de veículos em manutenção com base nas datas de entrada/saída.

ESTRUTURA DAS ABAS
------------------
- RESUMO: Carregada (não utilizada diretamente).
- CONTROLE_RENTALS: Recebe os dados reescritos.
- DB_GERAL: Contém a data de referência (A3).
- DB_ROTAS: Origem das datas e placas.
- DB_MANU: Registros de manutenção (placa, datas de entrada/saída).
- DB_PLACAS: Carregada (não utilizada diretamente).

FLUXO RESUMIDO
--------------
1. Abre o arquivo duas vezes:
   - Uma com valores calculados.
   - Outra para manter fórmulas e sobrescrever dados.
2. Lê a data de referência e valida formato.
3. Extrai datas e placas de DB_ROTAS até encontrar linha vazia.
4. Converte datas para formato brasileiro.
5. Preenche CONTROLE_RENTALS (colunas 2, 3 e 9).
6. Valida se a placa está em manutenção.
7. Salva alterações mantendo fórmulas.

REQUISITOS
----------
- Python 3.x
- Biblioteca:
    pip install openpyxl

COMO USAR
---------
1. Coloque "Planilha.xlsx" na pasta "Planilha/".
2. Execute:
    python script.py
3. Ao final, será exibida a mensagem:
    Seu programa foi concluído com sucesso!✅

MELHORIAS FUTURAS
-----------------
- Remover automaticamente traços ("-") das placas.
- Criar logs detalhados de execução.
- Permitir parametrização de caminhos e arquivos.

OBSERVAÇÃO IMPORTANTE
---------------------
A planilha utilizada neste script não está incluída neste repositório por conter dados internos e confidenciais do trabalho. 
Para executar o programa, substitua pelo seu próprio arquivo de planilha seguindo a estrutura de abas mencionada neste documento.

