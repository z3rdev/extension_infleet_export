# Infleet Ocorrências Exportador

Extensão local para exportar a lista de ocorrências em `https://app.infleet.com.br/occurrences` para Excel.

## Instalar no Chrome ou Edge

1. Abra `chrome://extensions` ou `edge://extensions`.
2. Ative o modo de desenvolvedor.
3. Clique em `Carregar sem compactação`.
4. Selecione esta pasta: `Raiz`.
5. Abra o Infleet, faca login manualmente e acesse `https://app.infleet.com.br/occurrences`.
6. Recarregue a pagina depois de instalar a extensão.

## Usar

1. Abra a extensão pelo icone do navegador.
2. Confira o periodo detectado na tela.
3. Se quiser, informe data inicial e final e clique em `Aplicar periodo`.
4. Clique em `Exportar Excel`.

A extensão muda o filtro `Mês atual` para `Personalizado` antes de preencher o período. Quando o Infleet abre o calendário `react-datepicker`, a extensão navega para o mês correto e clica diretamente nos dias do período. Durante a extração, um overlay bloqueia a tela para evitar scroll ou cliques manuais. A exportação rola a lista automaticamente e de-duplica as ocorrências pelo link de detalhe `/occurrences/{id}/details`. O arquivo gerado inclui data da ocorrência pelo agrupamento da lista, aba visualizada, tempo exibido, tipo de ocorrência, severidade, motorista, veiculo, periodo, ID e link.

Se o seletor de data do Infleet voltar para um periodo padrão, como `21/07`, a extensão ainda filtra o Excel pelo período informado no popup. Exemplo: com inicio em `01/08/2026`, linhas de julho não entram no arquivo final.

## Observação

A extensão nao armazena login nem senha. O preenchimento do periodo depende do seletor de datas renderizado pelo Infleet; se o site mudar esse componente, defina o periodo manualmente no Infleet e use apenas `Exportar Excel`.
