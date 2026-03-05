const axios = require('axios');
const https = require('https');
const fs = require('fs');
const zlib = require('zlib');
const listaClientes = require('./config_clientes'); // Importa sua lista

function salvarXml(conteudoBase64, chaveAcesso, nomeCliente) {
    try {
        const buffer = zlib.gunzipSync(Buffer.from(conteudoBase64, 'base64'));
        const pasta = `./notas_baixadas/${nomeCliente}`;
        
        if (!fs.existsSync(pasta)) fs.mkdirSync(pasta, { recursive: true });

        fs.writeFileSync(`${pasta}/${chaveAcesso}.xml`, buffer);
        console.log(`   💾 [${nomeCliente}] XML salvo.`);
    } catch (e) {
        console.error(`   ❌ Erro ao salvar XML: ${e.message}`);
    }
}

async function processarTodosOsClientes() {
    console.log(`🚀 Iniciando processamento de ${listaClientes.length} clientes...\n`);

    for (const cliente of listaClientes) {
        try {
            console.log(`📡 Conectando: ${cliente.nome} (${cliente.cnpj})`);

            // Criar agente HTTPS específico para este certificado
            const agente = new https.Agent({
                pfx: fs.readFileSync(cliente.pfx),
                passphrase: cliente.senha,
                rejectUnauthorized: true
            });

            const api = axios.create({ httpsAgent: agente });
            
            // URL de Produção Nacional para o dono do certificado
            const url = `https://adn.nfse.gov.br/contribuintes/DFe/0`;

            const resposta = await api.get(url);

            if (resposta.data.StatusProcessamento === "DOCUMENTOS_LOCALIZADOS") {
                const notas = resposta.data.LoteDFe;
                console.log(`   ✅ ${notas.length} notas encontradas.`);
                
                notas.forEach(nota => salvarXml(nota.ArquivoXml, nota.ChaveAcesso, cliente.nome));
            } else {
                console.log(`   ℹ️ Nenhuma nota nova.`);
            }

        } catch (erro) {
            console.error(`   ⚠️ Erro no cliente ${cliente.nome}: ${erro.response ? "Falha na API" : erro.message}`);
        }
        console.log("--------------------------------------------------");
    }
    console.log("🏁 Fim do processamento.");
}

processarTodosOsClientes();