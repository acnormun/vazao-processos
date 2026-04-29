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
