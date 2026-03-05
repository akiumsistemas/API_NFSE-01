/**
 * SISTEMA DE CAPTURA NFS-E NACIONAL - MB CONTABILIDADE
 * Versão: 2.8 - Iuri
 * * ALTERAÇÕES v2.8:
 * - PRESTADOR/TOMADOR VIA CHAVE: Utiliza a Chave de Acesso (índices 9 a 23) para extrair o CNPJ do emissor de forma direta e infalível.
 * - COMPETÊNCIA VIA XML: Mantém a extração profunda da tag <dCompet> para garantir a filtragem pelo mês correto (fato gerador).
 */

require('dotenv').config();
const axios = require('axios');
const https = require('https');
const fs = require('fs');
const zlib = require('zlib');
const readline = require('readline-sync');
const { XMLParser } = require('fast-xml-parser');
const listaClientes = require('./config_clientes');

const ARQUIVO_CONTROLE = './controle_nsu.json';
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const xmlParser = new XMLParser({ 
    ignoreAttributes: false, 
    removeNSPrefix: true 
});

function limparCnpj(cnpj) {
    return cnpj.toString().replace(/\D/g, '');
}

/**
 * Vasculha o XML atrás da data de competência
 */
function buscarNoObjeto(obj, chaveAlvo) {
    if (obj && obj[chaveAlvo]) return obj[chaveAlvo];
    for (let k in obj) {
        if (typeof obj[k] === 'object') {
            let found = buscarNoObjeto(obj[k], chaveAlvo);
            if (found) return found;
        }
    }
    return null;
}

function carregarControleNSU() {
    if (fs.existsSync(ARQUIVO_CONTROLE)) {
        return JSON.parse(fs.readFileSync(ARQUIVO_CONTROLE, 'utf-8'));
    }
    return {};
}

function salvarControleNSU(controle) {
    fs.writeFileSync(ARQUIVO_CONTROLE, JSON.stringify(controle, null, 2));
}

// --- NÚCLEO DE PROCESSAMENTO ---

function processarNota(nota, cliente, mesBusca, anoBusca) {
    try {
        const buffer = zlib.gunzipSync(Buffer.from(nota.ArquivoXml, 'base64'));
        const xmlString = buffer.toString('utf-8');
        const xmlObj = xmlParser.parse(xmlString);

        // 1. Busca Profunda apenas pela Competência (dCompet)
        const dCompet = buscarNoObjeto(xmlObj, 'dCompet');
        if (!dCompet) return false;

        const partesData = dCompet.split('-'); // Esperado: YYYY-MM-DD
        const anoNota = partesData[0];
        const mesNota = partesData[1];

        // Se a competência bater com o mês/ano solicitado...
        if (anoNota === anoBusca && mesNota === mesBusca) {
            
            const cnpjCliente = limparCnpj(cliente.cnpj);
            
            // 2. Extrai o CNPJ do Emitente direto da Chave de Acesso (Índices 9 a 23)
            const cnpjEmitenteNaChave = nota.ChaveAcesso.substring(9, 23);
            
            // 3. Define a pasta de forma simples e direta
            const tipo = (cnpjEmitenteNaChave === cnpjCliente) ? 'Prestadas' : 'Tomadas';

            const pastaFinal = `./notas_baixadas/${cliente.nome}/${mesNota}-${anoNota}/${tipo}`;
            if (!fs.existsSync(pastaFinal)) fs.mkdirSync(pastaFinal, { recursive: true });

            fs.writeFileSync(`${pastaFinal}/${nota.ChaveAcesso}.xml`, buffer);
            console.log(`      📥 [${tipo}] Salva (Comp: ${dCompet}) - Final da Chave: ...${nota.ChaveAcesso.substring(35)}`);
            return true;
        }
        return false;
    } catch (e) {
        // Silencia erros de estrutura de notas irrelevantes
        return false;
    }
}

// --- LOOP DE CAPTURA ---

async function iniciarCaptura() {
    console.log("====================================================");
    console.log("   SISTEMA DE CAPTURA NFS-E NACIONAL v2.8 - MB");
    console.log("====================================================\n");

    const competencia = readline.question("Digite o mes/ano de COMPETENCIA (Ex: 01/2026): ");
    const [mesAlvo, anoAlvo] = competencia.split('/');
    
    let controleNSU = carregarControleNSU();

    for (const cliente of listaClientes) {
        console.log(`\n▶️ Cliente: ${cliente.nome}`);
        
        if (!controleNSU[cliente.cnpj]) controleNSU[cliente.cnpj] = 0;
        
        let nsuAtual = controleNSU[cliente.cnpj];
        let interromperBusca = false;
        let totalCapturado = 0;

        try {
            const agente = new https.Agent({
                pfx: fs.readFileSync(cliente.pfx),
                passphrase: cliente.senha,
                rejectUnauthorized: true
            });

            const api = axios.create({ httpsAgent: agente });
            
            while (!interromperBusca) {
                await sleep(1500); 
                const url = `https://adn.nfse.gov.br/contribuintes/DFe/${nsuAtual}`;
                
                try {
                    const resposta = await api.get(url);
                    const dados = resposta.data;

                    if (dados.StatusProcessamento === "DOCUMENTOS_LOCALIZADOS" && dados.LoteDFe.length > 0) {
                        for (const nota of dados.LoteDFe) {
                            if (processarNota(nota, cliente, mesAlvo, anoAlvo)) {
                                totalCapturado++;
                            }
                            nsuAtual = nota.NSU; // Avança a fila
                        }
                    } else {
                        interromperBusca = true; // Fim dos documentos
                    }
                } catch (apiErro) {
                    if (apiErro.response && apiErro.response.status === 404) {
                        interromperBusca = true; // 404 = Fim da fila
                    } else { throw apiErro; }
                }
            }

            controleNSU[cliente.cnpj] = nsuAtual;
            salvarControleNSU(controleNSU);
            console.log(`   ✅ Sincronização Finalizada. Capturadas: ${totalCapturado}`);

        } catch (erro) {
            console.error(`   ❌ Falha Crítica no cliente ${cliente.nome}:`, erro.message);
        }
    }
    console.log("\n====================================================");
    console.log("   PROCESSO FINALIZADO - MB CONTABILIDADE");
    console.log("====================================================");
}

iniciarCaptura();