# 📄 Robô de Captura NFS-e Nacional (ADN) - MB Contabilidade

Este projeto é uma ferramenta de automação em Node.js desenvolvida para facilitar a rotina administrativa e fiscal da MB Contabilidade. O sistema conecta-se ao Ambiente de Dados Nacional (ADN) do Governo Federal, faz o download das Notas Fiscais de Serviço Eletrônicas (NFS-e), filtra por mês de competência e organiza os arquivos XML automaticamente.

## ✨ Funcionalidades

* **Multi-Empresas:** Processa a captura de múltiplos clientes em lote através de um arquivo de configuração.
* **Filtro por Competência Real:** Abre o XML do governo em tempo de execução e lê a tag `<dCompet>`, garantindo que a nota pertença ao mês do fato gerador (e não apenas à data de emissão).
* **Separação Automática:** Extrai o CNPJ diretamente da Chave de Acesso da nota para identificar com precisão infalível se o serviço foi **Prestado** (emitido pelo cliente) ou **Tomado** (recebido pelo cliente).
* **Controle Inteligente de Fila (NSU):** Registra o último número processado de cada cliente (`controle_nsu.json`) para nunca baixar a mesma nota duas vezes e economizar tempo nas consultas diárias.
* **Resiliência (Anti-Bloqueio):** Lida automaticamente com os limites de requisição do Governo (Erro 429) e identifica corretamente o fim da fila de notas (Erro 404).

---

## 🛠️ Pré-requisitos e Instalação do Node.js

Para que o robô funcione, o computador precisa ter o **Node.js** instalado (que é o "motor" que roda o nosso código). Se você ou alguém da equipe ainda não tem, siga os passos abaixo:

### Passo a Passo para Instalar o Node.js (Windows)
1. Acesse o site oficial: [https://nodejs.org/pt-br](https://nodejs.org/pt-br)

2. Você verá dois botões grandes. Clique no botão que diz **"LTS"** (Recomendado para a maioria dos usuários). Isso fará o download do instalador.

3. Abra o arquivo baixado e siga o padrão de instalação do Windows: clique em **Next** (Avançar), aceite os termos e continue clicando em **Next** até **Install** (Instalar) e, por fim, **Finish** (Concluir). *Não é necessário alterar nenhuma configuração padrão.*

4. **Verificando a instalação:** * Abra o terminal do seu computador (pode ser o `Prompt de Comando`, `PowerShell` ou o terminal do seu editor de código).
   * Digite `node -v` e aperte Enter. Se aparecer um número de versão (ex: `v20.x.x`), a instalação foi um sucesso!

---

## 📦 Instalação do Projeto

Com o Node.js instalado, você precisa instalar as "peças" (bibliotecas) que o nosso robô usa para funcionar.

1. Clone ou baixe a pasta deste projeto para o seu computador.
2. Abra o terminal **dentro da pasta do projeto** (onde está o arquivo `index.js`).
3. Execute o comando abaixo para instalar as dependências:

```bash
npm install axios dotenv fast-xml-parser readline-sync
⚙️ Configuração dos Clientes
Antes de rodar o robô, você precisa configurar os clientes que serão processados e fornecer os certificados digitais.

Crie uma pasta chamada certificados na raiz do projeto e coloque os arquivos .pfx dos clientes lá dentro.

Abra o arquivo config_clientes.js e preencha com os dados dos seus clientes seguindo este formato exato (sem pontos ou traços no CNPJ):

JavaScript
// config_clientes.js
module.exports = [
    {
        nome: "EMPRESA EXEMPLO LTDA",
        cnpj: "12345678000100",
        pfx: "./certificados/empresa_exemplo.pfx",
        senha: "senha_do_certificado"
    },
    // Adicione os demais clientes aqui...
];
🚀 Como Usar
Devido aos padrões de criptografia exigidos pelos certificados A1 brasileiros, é obrigatório executar o Node.js informando que ele deve aceitar o padrão antigo de segurança (flag legacy).

No terminal, dentro da pasta do projeto, execute o comando:

Bash
node --openssl-legacy-provider index.js
O que vai acontecer na tela:
O sistema solicitará que você digite a Competência desejada (Exemplo: 02/2026).

O robô vai se conectar ao Governo para o primeiro cliente da lista.

Ele vai baixar a nota, ler o XML e verificar:

É da competência informada? (Se sim, ele guarda; se não, ele descarta).

O CNPJ da chave é igual ao do cliente? (Se sim, vai para a pasta Prestadas; se não, vai para Tomadas).

Os arquivos serão salvos organizados na estrutura:
./notas_baixadas/[Nome do Cliente]/[Mes-Ano]/[Prestadas ou Tomadas]/

🧠 Como funciona a Lógica por trás do Robô (Para Suporte)
Para fins de suporte e manutenção, é importante entender como o robô toma decisões:

Fila NSU (Número Sequencial Único): O governo não permite pesquisar por data, ele entrega as notas em uma "fila" contínua. O arquivo controle_nsu.json anota em qual número de nota a leitura parou para cada cliente. Na próxima vez que o robô rodar, ele continua exatamente de onde parou.

Dica de Manutenção (Reset do Robô): Se você precisar forçar o robô a reler todas as notas de um cliente do zero (caso ache que alguma nota retroativa não foi lida), basta apagar o arquivo controle_nsu.json antes de executar o comando. O sistema recriará o arquivo e varrerá tudo novamente, salvando apenas as notas do mês que você digitar.

Identificação do Prestador: A Chave de Acesso da NFS-e possui 50 dígitos. O CNPJ de quem emitiu a nota fica posicionado sempre do 10º ao 23º dígito. O robô recorta essa informação da chave e compara com o CNPJ do cadastro para definir se a nota vai para a pasta de prestadas ou tomadas de forma matemática e sem erros.