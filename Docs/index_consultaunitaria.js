require('dotenv').config();
const axios = require('axios');
const https = require('https');
const fs = require('fs');
const zlib = require('zlib');

/**
 * Função para converter o conteúdo do Governo (Base64 + GZIP) em arquivo XML real
 */
function salvarXml(conteudoBase64, chaveAcesso) {
    try {
        // 1. Converte Base64 para binário (Buffer)
        const bufferCompactado = Buffer.from(conteudoBase64, 'base64');

        // 2. Descompacta o GZIP (padrão do portal nacional)
        const xmlDescompactado = zlib.gunzipSync(bufferCompactado);

        // 3. Define o nome do arquivo e salva na pasta do projeto
        const nomeArquivo = `./${chaveAcesso}.xml`;
        fs.writeFileSync(nomeArquivo, xmlDescompactado);

        console.log(`   💾 Arquivo salvo com sucesso: ${nomeArquivo}`);
    } catch (erro) {
        console.error(`   ❌ Erro ao processar nota ${chaveAcesso}:`, erro.message);
    }
}

/**
 * Função principal de conexão e busca
 */
async function buscarNotasProducao() {
    try {
        console.log("--- Iniciando Busca no Portal Nacional (PRODUÇÃO) ---");

        // Verificação de segurança das variáveis
        if (!process.env.CAMINHO_CERTIFICADO || !process.env.URL_API_PRODUCAO) {
            throw new Error("Configurações ausentes no arquivo .env");
        }

        // 1. Carregar o Certificado Digital A1
        const certificadoBuffer = fs.readFileSync(process.env.CAMINHO_CERTIFICADO);

        // 2. Configurar o Agente HTTPS com o certificado
        const agenteSeguranca = new https.Agent({
            pfx: certificadoBuffer,
            passphrase: process.env.SENHA_CERTIFICADO,
            // Em produção, o ideal é manter a verificação de certificados ativa (true)
            rejectUnauthorized: true 
        });

        // 3. Criar instância do Axios com o agente configurado
        const api = axios.create({ httpsAgent: agenteSeguranca });

        console.log("📡 Conectando ao servidor do Governo...");
        const resposta = await api.get(process.env.URL_API_PRODUCAO);

        // 4. Analisar o resultado
        const dados = resposta.data;

        if (dados.StatusProcessamento === "DOCUMENTOS_LOCALIZADOS" && dados.LoteDFe.length > 0) {
            console.log(`✅ Sucesso! ${dados.LoteDFe.length} documentos encontrados.`);
            
            // Percorre cada nota encontrada no lote
            dados.LoteDFe.forEach((item, index) => {
                console.log(`\n[Nota ${index + 1}] Chave: ${item.ChaveAcesso}`);
                
                if (item.ArquivoXml) {
                    salvarXml(item.ArquivoXml, item.ChaveAcesso);
                } else {
                    console.log("   ⚠️ Nota sem conteúdo XML disponível no lote.");
                }
            });

            console.log("\n--- Processamento Finalizado ---");
        } else {
            console.log("ℹ️ Nenhuma nota nova localizada para este CNPJ no momento.");
        }

    } catch (erro) {
        if (erro.response) {
            console.error("\n❌ Erro retornado pelo Governo:");
            console.error(`Status: ${erro.response.status}`);
            console.dir(erro.response.data, { depth: null });
        } else {
            console.error("\n❌ Erro técnico de configuração:", erro.message);
        }
    }
}

// Executa a função
buscarNotasProducao();