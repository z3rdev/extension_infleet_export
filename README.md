# Infleet Ocorrencias Exportador

Extensao local para exportar a lista de ocorrencias em `https://app.infleet.com.br/occurrences` para Excel.

## Instalar no Chrome ou Edge

1. Abra `chrome://extensions` ou `edge://extensions`.
2. Ative o modo de desenvolvedor.
3. Clique em `Carregar sem compactacao`.
4. Selecione esta pasta: `E:\PROJETOS\EXTENSÃO - INFLEET EXTRAÇÃO DE OCORRENCIA`.
5. Abra o Infleet, faca login manualmente e acesse `https://app.infleet.com.br/occurrences`.
6. Recarregue a pagina depois de instalar a extensao.

## Usar

1. Abra a extensao pelo icone do navegador.
2. Confira o periodo detectado na tela.
3. Se quiser, informe data inicial e final e clique em `Aplicar periodo`.
4. Clique em `Exportar Excel`.

A extensao muda o filtro `Mes atual` para `Personalizado` antes de preencher o periodo. Quando o Infleet abre o calendario `react-datepicker`, a extensao navega para o mes correto e clica diretamente nos dias do periodo. Durante a extracao, um overlay bloqueia a tela para evitar scroll ou cliques manuais. A exportacao rola a lista automaticamente e de-duplica as ocorrencias pelo link de detalhe `/occurrences/{id}/details`. O arquivo gerado inclui data da ocorrencia pelo agrupamento da lista, aba visualizada, tempo exibido, tipo de ocorrencia, severidade, motorista, veiculo, periodo, ID e link.

Se o seletor de data do Infleet voltar para um periodo padrao, como `21/07`, a extensao ainda filtra o Excel pelo periodo informado no popup. Exemplo: com inicio em `01/08/2026`, linhas de julho nao entram no arquivo final.

## Observacao

A extensao nao armazena login nem senha. O preenchimento do periodo depende do seletor de datas renderizado pelo Infleet; se o site mudar esse componente, defina o periodo manualmente no Infleet e use apenas `Exportar Excel`.
