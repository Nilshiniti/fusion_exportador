# Exportador do Fusion

Script para o Autodesk Fusion que baixa seus arquivos da nuvem para o computador, em lote: F3D/F3Z, STEP e, se quiser, STL.

O script só lê da nuvem. Ele abre cada arquivo e fecha sem salvar; não salva, apaga, move nem renomeia nada na nuvem.

Projeto independente, sem vínculo com a Autodesk.

## O que ele faz

- Janela para escolher o hub, os projetos e os formatos.
- Um erro em um arquivo vai para o registro e a exportação continua.
- Rodar de novo na mesma pasta retoma de onde parou, pulando o que já saiu.
- Mantém acentos e hífens nos nomes, trocando só os caracteres proibidos no Windows.
- Nunca sobrescreve: nomes parecidos recebem `(2)`, e versões novas recebem `v5`.
- Avisa antes de começar se algum caminho for passar do limite do Windows.
- Barra de progresso com cancelamento seguro e resumo no final.

## Instalação

Não é preciso instalar programa algum: o script roda dentro do Fusion.

1. Na página do repositório, clique em **Code > Download ZIP** e extraia o arquivo.
2. **Renomeie a pasta extraída para `fusion_exportador`** (ela vem como `fusion_exportador-main`). O nome da pasta precisa ser igual ao do `.py`, senão o Fusion não reconhece o script.
3. Copie a pasta para:
   - **Windows:** `%APPDATA%\Autodesk\Autodesk Fusion 360\API\Scripts`
   - **Mac:** use o botão **+** em **Scripts e Suplementos** para adicionar a pasta. Não testado.
4. No Fusion, abra **Utilitários > Scripts e Suplementos** (atalho: `Shift+S`) e localize `fusion_exportador`. Se ele não aparecer, feche e abra o Fusion.

## Como usar

1. Salve e feche os arquivos em que estiver trabalhando.
2. Crie uma pasta vazia, de caminho curto, para receber os arquivos.
3. Em **Scripts e Suplementos**, selecione `fusion_exportador` e clique em **Executar**.
4. Na janela, escolha o hub, os projetos e os formatos, e clique em **Continuar**. No primeiro uso, comece por um projeto pequeno.
5. Escolha a pasta de destino.
6. Confira as quantidades no aviso e responda **Sim**.
7. Aguarde o resumo final. Os arquivos abrem e fecham sozinhos na tela, então evite usar o Fusion nesse período.

**Cancelar** termina o que estiver em andamento, fecha o arquivo e para. Para continuar depois, rode de novo e escolha a mesma pasta.

Durante exportações longas, mantenha o computador na tomada e sem suspender ou hibernar. O script não impede a suspensão.

## Onde ficam os arquivos

```
<destino>\<hub>\<projeto>\<pastas da nuvem>\<arquivo>.f3d
<destino>\<hub>\<projeto>\<pastas da nuvem>\<arquivo>.step
<destino>\_exportador\registro.csv          uma linha por tentativa e formato
<destino>\_exportador\log_<data_hora>.txt   detalhes, inclusive erros
```

Não apague o `registro.csv`: é ele que permite retomar. Sem ele, o script não sabe o que já saiu e, como não sobrescreve nada, cria cópias com `(2)`.

O CSV usa `;` como separador e abre direto no Excel em português.

## Como ele se comporta

| Situação | O que acontece |
|---|---|
| Erro em um arquivo | Vai para o CSV e o log, e a exportação segue. |
| Rodar de novo na mesma pasta | Pula o que já saiu e tenta de novo o que falhou. |
| Apagou um arquivo exportado | É refeito com o mesmo nome. |
| Versão nova na nuvem | Salva `nome v5.f3d` e mantém o anterior. |
| Nomes parecidos (`A:B` e `A_B`) | O segundo vira `A_B (2)`. |
| Arquivo já aberto com alterações não salvas | Não é fechado nem exportado. Fica como erro até você salvar ou fechar. |
| Fusion fecha no meio | Na próxima execução, esse arquivo é pulado uma vez e registrado; na seguinte, é tentado de novo. |
| `registro.csv` aberto no Excel | No início, o script pede para fechar e encerra. No meio, guarda as linhas e pede para fechar no final. |
| Caminho acima de 250 caracteres | Avisa antes de começar. Se continuar, esses arquivos ficam como erro; para incluí-los, use um destino mais curto. |
| Desenhos (F2D), PDFs e outros tipos | Não são baixados. Ficam como `pulado` no CSV, com o tipo do arquivo. |

Os limites de tamanho de nomes e caminhos ficam nas primeiras linhas do `.py`, em "Ajustes".

## O que já foi testado

**No Fusion real** (versão 2705.1.30, Windows 11, Python 3.14): uma execução exportou 192 arquivos em STEP, sem erros, e pulou 2 que não eram modelos (um PDF e um desenho).

**Só em simulação, ainda não no Fusion real:**

- exportação em F3D, F3Z e STL;
- retomada em uma segunda execução;
- cancelamento pelo botão;
- comportamento com arquivos já abertos, caminhos longos e queda do Fusion;
- macOS.

Se encontrar um problema, abra uma issue com o `log_<data_hora>.txt`. Antes de enviar, confira o conteúdo: o log e o CSV trazem nomes de hubs, projetos e arquivos.

## Limitações

- Baixa só modelos do Fusion. Desenhos (F2D), PDFs e outros tipos não são baixados.
- Exporta a versão mais recente de cada arquivo.
- Avisos do Fusion ao abrir um arquivo, como referência ausente, podem parar a exportação até você responder.
- Requer conexão com a internet.

## Licença

[MIT](LICENSE). O software é fornecido sem garantia; comece por um projeto pequeno.
