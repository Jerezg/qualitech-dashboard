# Qualitech — Dashboard de Performance Operacional (app Streamlit)

Aplicação web que lê o arquivo `Performance_<Mês>.xlsx` diretamente da pasta
sincronizada do OneDrive/SharePoint e recalcula todos os indicadores e
gráficos sozinha, sempre que o arquivo é atualizado. Não é mais necessário
pedir para o Claude gerar um HTML novo a cada revisão — basta salvar a
planilha atualizada na pasta de sempre e clicar em **"🔄 Atualizar agora"**
(ou deixar o "Auto (60s)" ligado).

## 1. Instalação (uma única vez)

Requer Python 3.10+ instalado na máquina — no instalador do Python
(python.org), marque a caixinha **"Add python.exe to PATH"** antes de
clicar em Install (sem isso, o Windows não sabe onde achar o Python).

**Opção mais fácil (Windows):** nem precisa instalar as dependências à
mão — pule direto para a seção 2 e dê duplo clique em
`Iniciar_Dashboard.bat`, ele instala tudo sozinho na primeira vez.

**Ou, manualmente, via terminal:**

```bash
cd qualitech_app
pip install -r requirements.txt
```

## 2. Rodar o app

**Opção mais fácil (Windows):** dê duplo clique em **`Iniciar_Dashboard.bat`**,
dentro desta mesma pasta. Ele confere/instala as dependências sozinho e
inicia o dashboard — não precisa digitar nenhum comando.

**Ou, manualmente, via terminal:**

```bash
streamlit run app.py
```

> Se aparecer o erro **"'streamlit' não é reconhecido como um comando..."**,
> é porque a pasta de scripts do Python não está na PATH do Windows. Use
> `python -m streamlit run app.py` em vez de `streamlit run app.py`
> (funciona mesmo sem isso) — ou simplesmente use o
> `Iniciar_Dashboard.bat`, que já contorna esse problema sozinho.
>
> Se aparecer um erro **"OSError: [Errno 2] No such file or directory"**
> com um caminho gigante cheio de `AppData\Local\Packages\...`, o
> problema é o **Python da Microsoft Store** — essa versão instala tudo
> numa pasta tão profunda que esbarra no limite de 260 caracteres do
> Windows para caminhos de arquivo, e a instalação de pacotes falha no
> meio do caminho (literalmente). A solução é trocar pelo Python
> "de verdade": desinstale o "Python" da Microsoft Store (Configurações
> → Aplicativos → procure "Python" → Desinstalar) e instale o oficial em
> [python.org/downloads](https://www.python.org/downloads/), marcando
> **"Add python.exe to PATH"** na instalação. O `Iniciar_Dashboard.bat`
> já detecta esse caso e avisa na tela se identificar o Python da Store.

Isso abre automaticamente uma aba no navegador em `http://localhost:8501`.

⚠️ **Atenção — a janela preta (terminal/"DOS") precisa ficar aberta o tempo
todo.** Ela é o "motor" que mantém o dashboard rodando — a aba do navegador
é só a "tela"; se você fechar a janela preta (ou clicar no X dela), o
dashboard para na hora, mesmo que a aba do navegador continue aberta (ela
vai mostrar erro de conexão). Pode minimizar a janela preta à vontade,
só não feche. Para parar de verdade, é só fechar essa janela normalmente
(ou apertar `Ctrl+C` dentro dela) quando terminar de usar.

Se isso incomodar no dia a dia, existe a opção de publicar o dashboard num
link fixo (ex.: Streamlit Community Cloud) que fica sempre no ar sem
precisar de nenhuma janela aberta na sua máquina — veja a seção 8.

## 3. Base de Recursos (RH) — opcional

Se o arquivo `Planejamento_Ativos_Status_*.xlsx` (efetivo, turnover, novos
admitidos, bloqueios, certificações) estiver na mesma pasta sincronizada
configurada acima, o app já encontra ele sozinho (procura por "ativos" ou
"planejamento" no nome do arquivo) e libera a seção **"Recursos & Pessoas"**,
com indicadores de headcount que complementam os de diárias/WO. Sem esse
arquivo, essa seção fica desabilitada e o resto do dashboard funciona
normalmente.

## 4. Apontar para os seus dados

Na barra lateral, o campo **"Pasta sincronizada (OneDrive/SharePoint)"** já
vem preenchido com o caminho da pasta de Agosto/2026:

```
C:\Users\Jorge Gonçalves\OneDrive - Qualitech Inspeção, Reparo e Manutenção Ltda\Área de Trabalho\WIP\9. Performance\Ago-2026
```

- Se você apontar para a pasta **"9. Performance"** (a pasta-mãe, sem
  especificar o mês), o app lista automaticamente todos os `.xlsx`
  encontrados em qualquer subpasta de mês e deixa você escolher qual abrir.
- Se o arquivo `.xlsx` estiver aberto no Excel ao mesmo tempo, normalmente
  ainda é possível ler (modo somente leitura). Se der erro de arquivo
  bloqueado, salve e feche o Excel antes de clicar em "Atualizar agora".
- Caso a pasta não seja encontrada (por exemplo, rodando em outra máquina ou
  sem o OneDrive sincronizado), o app mostra um botão para enviar o `.xlsx`
  manualmente, como alternativa.

## 5. Atualização automática

- **🔄 Atualizar agora**: força a releitura imediata do arquivo (limpa o
  cache).
- **Auto (60s)**: religa a leitura automaticamente a cada 60 segundos,
  útil para deixar o dashboard aberto num monitor/TV da sala de reunião.
- Mesmo sem clicar em nada, sempre que a página é recarregada o app compara
  a data de modificação do arquivo (`mtime`) e релê os dados se algo mudou.

## 6. Filtros

Na barra lateral, em **Filtros**, é possível recortar os indicadores por
Coordenador, Cliente, Tipo de Contrato (Fixa/Variável/Spot) e Tipo de
Serviço. Os filtros afetam as diárias, ordens de serviço e os gráficos/
tabelas de composição; os 4 indicadores de topo (Man-days Realizados,
Utilização, Backlog, SISPAT) refletem sempre o total consolidado do mês,
por virem de abas-resumo da planilha.

## 7. Estrutura do projeto

```
qualitech_app/
├── app.py                  → aplicação Streamlit (layout, gráficos, filtros)
├── data.py                 → leitura e cálculo dos indicadores de diárias/WO
├── workforce.py            → leitura e cálculo dos indicadores de RH/efetivo
├── Iniciar_Dashboard.bat   → atalho p/ Windows: dê duplo clique p/ rodar o App
├── requirements.txt        → dependências Python
├── .streamlit/
│   └── config.toml         → tema visual (cores da marca Qualitech)
└── README.md               → este arquivo
```

## 8. Próximo passo — deixar acessível para outras pessoas (opcional)

Hoje o app roda localmente (`streamlit run app.py`), visível só na sua
máquina. Se quiser que outras pessoas da diretoria acessem por um link,
sem precisar instalar nada, existem 3 caminhos — depois de validar que o
layout e os números estão do jeito que você quer, é só me avisar qual
prefere que eu preparo:

1. **Servidor interno da Qualitech** (mais controle, dados nunca saem da
   rede da empresa) — roda o mesmo `streamlit run app.py` num servidor/VM
   interno, com acesso via IP ou nome interno.
2. **Streamlit Community Cloud** (gratuito, mais rápido de colocar no ar) —
   exige que o arquivo de dados fique num repositório ou storage que o
   Streamlit Cloud consiga acessar (não lê diretamente uma pasta do seu
   OneDrive local).
3. **Nuvem própria da Qualitech** (Azure, por já usarem OneDrive/SharePoint
   da Microsoft) — permite ler o Excel direto do SharePoint via API, o mais
   robusto a médio prazo, porém com mais configuração inicial.

Qualquer uma dessas opções reaproveita o `app.py`/`data.py` já prontos —
só muda onde e como ele fica hospedado.
