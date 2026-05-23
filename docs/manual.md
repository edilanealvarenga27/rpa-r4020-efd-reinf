Manual de Uso – Robô R-4020
1. Objetivo

O Robô R-4020 foi desenvolvido para automatizar o preenchimento de informações no sistema EFD-Reinf por meio de automação de teclado e mouse.

O sistema realiza leitura de dados diretamente de uma planilha Excel e executa automaticamente os lançamentos no navegador.

2. Requisitos

Antes da execução, é necessário:

Windows 10 ou superior;
Google Chrome instalado;
acesso ao sistema EFD-Reinf;
planilha Excel preenchida;
resolução de tela padrão recomendada;
zoom do navegador em 100%.

3. Arquivos Necessários

O usuário deve possuir:

Robo_R4020.exe
Planilha Excel (.xlsx)

4. Como Executar
Passo 1 – Abrir o sistema
Acesse normalmente o sistema EFD-Reinf no navegador Google Chrome.
Navegue até a tela do evento R-4020.
Deixe a tela pronta para iniciar os lançamentos.
Passo 2 – Abrir o robô

Execute:

Robo_R4020.exe

A tela abaixo será exibida:

seleção da planilha;
aba da planilha;
linha inicial;
linha final;
configuração das colunas.

5. Seleção da Planilha

Clique em:

Procurar

Selecione a planilha Excel que contém os dados.

6. Configuração das Colunas

Informe as colunas correspondentes às informações da planilha.

Exemplo padrão:

Campo	Coluna
Período de Apuração	AG
Estabelecimento / CNPJ Unidade	C6
Beneficiário / Recolhedor DARF	L
Natureza do Rendimento	AP
Observação Pré-Doc	E
Data OB Pagamento	Z
Valor Doc de Origem	W
Base de Cálculo DARF	S
Valor Receita DARF	T

Observações:

É possível utilizar:
colunas (AG, L, T);
células fixas (C6).
As informações serão lidas automaticamente pelo robô.

7. Linha Inicial e Final

Informe:

linha inicial da planilha;
linha final da planilha.

Exemplo:

Linha inicial: 9
Linha final: 20

O robô processará todas as linhas dentro desse intervalo.

8. Iniciando a Execução

Clique em:

Iniciar Robô

O sistema solicitará o posicionamento do mouse sobre o botão:

R-4020

Após clicar em OK:

haverá contagem regressiva;
posicione o mouse sobre o botão R-4020;
o robô capturará automaticamente a posição.

9. Durante a Execução

Durante a execução:

não utilize teclado;
não mova o mouse;
não altere a janela do navegador;
não minimize o sistema;
aguarde o término do processamento.

A janela exibirá:

logs em tempo real;
etapas executadas;
progresso da automação;
tempos de execução.

10. Botão “Parar Robô”

Caso necessário, utilize:

Parar Robô

O sistema interromperá a execução de forma segura na próxima etapa disponível.

11. Logs

O robô gera automaticamente:

log_robo_r4020.txt

O arquivo contém:

etapas executadas;
tempos;
erros;
informações de debug.

12. Captura de Erros

Em caso de falha:

o robô gera automaticamente uma captura de tela;
o arquivo é salvo na mesma pasta do executável.

Exemplo:

erro_preenchimento_20260523_143210.png

13. Recomendações Importantes
Navegador

Utilizar preferencialmente:

Google Chrome

O Firefox pode apresentar diferenças na navegação por TAB.

Zoom do Navegador

Utilizar:

100%
Resolução

Preferencialmente:

manter resolução padrão;
evitar múltiplos monitores;
evitar mudanças durante a execução.

14. Problemas Comuns
O robô digitou no local errado

Possíveis causas:

mudança de foco da janela;
movimentação do mouse;
alteração da tela;
zoom diferente;
navegador diferente.
O executável não abre

Possíveis causas:

antivírus bloqueando;
permissão do Windows;
extração incompleta do ZIP/7z.

15. Observações Finais

O robô foi desenvolvido para automação operacional de lançamentos repetitivos, visando:

redução de tempo;
aumento de produtividade;
padronização das atividades;
diminuição de erros manuais.

Recomenda-se sempre:

validar os dados antes da execução;
acompanhar os primeiros lançamentos;
manter cópia da planilha original.
