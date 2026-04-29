# Vazao de Processos

Ferramenta simples para cruzar a planilha de distribuicoes/entradas com a planilha de saidas e calcular a vazao dos processos.

## Interface

```powershell
python .\interface_vazao.py
```

Selecione:

1. o CSV de distribuicoes/entradas;
2. o XLSX de saidas;
3. o local do arquivo de resultado.

Depois de gerar a planilha, a mesma tela mostra a vazao por assessor com filtros de tarefa e periodo:

- todos;
- este mes;
- ultimos 30 dias;
- este ano;
- ultimos 12 meses;
- personalizado, usando datas no formato `dd/mm/aaaa`.

A tabela exibe quantidade de entradas, quantidade de saidas, saldo, media/mediana/minimo/maximo de vazao em dias, entradas sem saida e saidas sem entrada.

Use o botao `Exportar PDF` para salvar o relatorio filtrado que esta na tela.

## Linha de comando

```powershell
python .\cruzar_vazao.py "C:\caminho\processos-triagem.csv" "C:\caminho\PJyVh.xlsx" -o "resultado.xlsx"
```

## Regras do cruzamento

- cruza por numero do processo;
- respeita a sequencia cronologica quando o processo entra, sai e volta;
- quando ha entrada e saida, calcula `Vazao (dias)`;
- quando entrou mas ainda nao saiu, deixa `Data saida` e `Vazao (dias)` vazias;
- quando saiu mas nao consta na entrada, deixa `Data entrada` e `Vazao (dias)` vazias;
- quando o assessor de entrada e saida divergem, usa o assessor da saida.
- nas saidas, considera apenas `Minutar relatório de voto` e `Minutar decisão monocrática`.
