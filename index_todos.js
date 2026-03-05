/**
 * SISTEMA DE CAPTURA NFS-E NACIONAL - MB CONTABILIDADE
 * Versão: 3.0
 * * ALTERAÇÕES v3.0:
 * - SINCRONIZAÇÃO CONTÍNUA: Baixa e organiza tudo, não descarta notas de outros meses.
 * - RELATÓRIO EXCEL: Gera arquivo .xlsx com Resumo, Alertas e Lista de Notas.
 * - AUDITORIA: Detecta numeração de notas Prestadas puladas (gaps).
 * - EXTRAÇÃO DE DADOS: Captura Número da Nota, Valor e Nomes dos Atores.
 */

require('dotenv').config();
const axios = require('axios');
const https = require('https');
const fs = require('fs');
const zlib = require('zlib');
const readline = require('readline-sync');
const { XMLParser } = require('fast-xml-parser');
const ExcelJS = require('exceljs'); // Nova biblioteca
const listaClientes = require('./config_clientes');

const ARQUIVO_CONTROLE = './controle_nsu.json';
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const xmlParser = new XMLParser({ ignoreAttributes: false, removeNSPrefix: true });

// --- FUNÇÕES AUXILIARES ---

function limparCnpj(cnpj) { return cnpj.toString().replace(/\D/g, ''); }

function buscarNoObjeto(obj, chaveAlvo) {
    if (obj && obj[chaveAlvo] !== undefined) return obj[chaveAlvo];
    for (let k in obj) {
        if (typeof obj[k] === 'object') {
            let found = buscarNoObjeto(obj[k], chaveAlvo);
            if (found !== null) return found;
        }
    }
    return null;
}

function carregarControleNSU() {
    if (fs.existsSync(ARQUIVO_CONTROLE)) return JSON.parse(fs.readFileSync(ARQUIVO_CONTROLE, 'utf-8'));
    return {};
}

function salvarControleNSU(controle) {
    fs.writeFileSync(ARQUIVO_CONTROLE, JSON.stringify(controle, null, 2));
}

// --- GERAÇÃO DO EXCEL ---

async function gerarRelatorioExcel(cliente, mesAlvo, anoAlvo, notasProcessadas, mesesExtras) {
    const workbook = new ExcelJS.Workbook();
    const abaResumo = workbook.addWorksheet('Resumo e Alertas');
    const abaNotas = workbook.addWorksheet('Lista de Notas');

    const notasDoMesAlvo = notasProcessadas.filter(n => n.mesAno === `${mesAlvo}/${anoAlvo}`);
    const prestadas = notasDoMesAlvo.filter(n => n.tipo === 'Prestadas');
    const tomadas = notasDoMesAlvo.filter(n => n.tipo === 'Tomadas');

    const totalValorPrestado = prestadas.reduce((acc, curr) => acc + curr.valor, 0);
    const totalValorTomado = tomadas.reduce((acc, curr) => acc + curr.valor, 0);

    // LÓGICA DE AUDITORIA DE NUMERAÇÃO (Apenas Prestadas)
    let numerosPulados = [];
    if (prestadas.length > 0) {
        const numeros = prestadas.map(n => parseInt(n.numero)).filter(n => !isNaN(n)).sort((a, b) => a - b);
        if (numeros.length > 0) {
            const min = numeros[0];
            const max = numeros[numeros.length - 1];
            for (let i = min; i <= max; i++) {
                if (!numeros.includes(i)) numerosPulados.push(i);
            }
        }
    }

    // --- CONSTRUÇÃO DA ABA RESUMO ---
    abaResumo.columns = [{ width: 40 }, { width: 40 }];
    abaResumo.addRow(['📊 RELATÓRIO DE CAPTURA - MB CONTABILIDADE', '']);
    abaResumo.addRow(['Cliente:', cliente.nome]);
    abaResumo.addRow(['Competência Alvo:', `${mesAlvo}/${anoAlvo}`]);
    abaResumo.addRow(['', '']);

    abaResumo.addRow(['RESUMO DO MÊS ALVO', '']);
    abaResumo.addRow(['Qtd. Notas Prestadas:', prestadas.length]);
    abaResumo.addRow(['Valor Total Prestadas:', `R$ ${totalValorPrestado.toFixed(2)}`]);
    abaResumo.addRow(['Qtd. Notas Tomadas:', tomadas.length]);
    abaResumo.addRow(['Valor Total Tomadas:', `R$ ${totalValorTomado.toFixed(2)}`]);
    abaResumo.addRow(['', '']);

    abaResumo.addRow(['⚠️ ALERTAS DO SISTEMA', '']);
    
    // Alerta de Gaps
    if (numerosPulados.length > 0) {
        abaResumo.addRow(['Notas Faltantes (Puladas):', numerosPulados.join(', ')]);
    } else if (prestadas.length > 0) {
        abaResumo.addRow(['Auditoria de Sequência:', 'Nenhuma nota pulada identificada.']);
    } else {
        abaResumo.addRow(['Auditoria de Sequência:', 'Sem notas prestadas para auditar.']);
    }

    // Alerta de Meses Extras
    const chavesMesesExtras = Object.keys(mesesExtras);
    if (chavesMesesExtras.length > 0) {
        abaResumo.addRow(['Notas de Outros Meses Capturadas:', 'O NSU trouxe notas antigas/futuras:']);
        chavesMesesExtras.forEach(m => {
            abaResumo.addRow([`-> Competência ${m}:`, `${mesesExtras[m]} nota(s) salva(s) na respectiva pasta.`]);
        });
    } else {
        abaResumo.addRow(['Notas de Outros Meses:', 'Nenhuma.']);
    }

    // --- CONSTRUÇÃO DA ABA NOTAS ---
    abaNotas.columns = [
        { header: 'Competência', key: 'comp', width: 15 },
        { header: 'Tipo', key: 'tipo', width: 15 },
        { header: 'Número (nNFSe)', key: 'num', width: 15 },
        { header: 'Valor (R$)', key: 'val', width: 15 },
        { header: 'Emitente', key: 'emit', width: 40 },
        { header: 'Tomador', key: 'toma', width: 40 },
        { header: 'Chave de Acesso', key: 'chave', width: 55 }
    ];

    // Adiciona todas as notas capturadas no Excel
    notasProcessadas.forEach(n => {
        abaNotas.addRow({
            comp: n.mesAno, tipo: n.tipo, num: n.numero, val: n.valor,
            emit: n.nomeEmitente, toma: n.nomeTomador, chave: n.chave
        });
    });

    const pastaRelatorio = `./notas_baixadas/${cliente.nome}/Relatorios`;
    if (!fs.existsSync(pastaRelatorio)) fs.mkdirSync(pastaRelatorio, { recursive: true });
    
    await workbook.xlsx.writeFile(`${pastaRelatorio}/Auditoria_${mesAlvo}-${anoAlvo}.xlsx`);
    
    // EXIBE ALERTAS NO TERMINAL
    if (numerosPulados.length > 0) console.log(`      ⚠️ ALERTA: Faltam as notas prestadas: ${numerosPulados.join(', ')}`);
    chavesMesesExtras.forEach(m => console.log(`      ⚠️ ALERTA: ${mesesExtras[m]} nota(s) encontrada(s) do mês ${m}`));
}

// --- PROCESSAMENTO E EXTRATOS ---

function processarNota(nota, cliente) {
    try {
        const buffer = zlib.gunzipSync(Buffer.from(nota.ArquivoXml, 'base64'));
        const xmlString = buffer.toString('utf-8');
        const xmlObj = xmlParser.parse(xmlString);

        const dCompet = buscarNoObjeto(xmlObj, 'dCompet');
        if (!dCompet) return null;

        const partesData = dCompet.split('-'); 
        const mesAnoFormatado = `${partesData[1]}/${partesData[0]}`; // MM/YYYY
        
        const cnpjCliente = limparCnpj(cliente.cnpj);
        const cnpjEmitenteNaChave = nota.ChaveAcesso.substring(9, 23);
        const tipo = (cnpjEmitenteNaChave === cnpjCliente) ? 'Prestadas' : 'Tomadas';

        // Extração de Dados Extras para o Excel
        const nNFSe = buscarNoObjeto(xmlObj, 'nNFSe') || 'N/A';
        const vServ = buscarNoObjeto(xmlObj, 'vServ') || buscarNoObjeto(xmlObj, 'vLiq') || 0;
        
        const emitObj = buscarNoObjeto(xmlObj, 'emit');
        const nomeEmitente = emitObj ? buscarNoObjeto(emitObj, 'xNome') : 'N/A';
        
        const tomaObj = buscarNoObjeto(xmlObj, 'toma') || buscarNoObjeto(xmlObj, 'tomaServico');
        const nomeTomador = tomaObj ? buscarNoObjeto(tomaObj, 'xNome') : 'N/A';

        // Salva o XML fisicamente
        const pastaFinal = `./notas_baixadas/${cliente.nome}/${partesData[1]}-${partesData[0]}/${tipo}`;
        if (!fs.existsSync(pastaFinal)) fs.mkdirSync(pastaFinal, { recursive: true });
        fs.writeFileSync(`${pastaFinal}/${nota.ChaveAcesso}.xml`, buffer);

        return {
            chave: nota.ChaveAcesso,
            numero: nNFSe,
            valor: parseFloat(vServ),
            tipo: tipo,
            mesAno: mesAnoFormatado,
            nomeEmitente: nomeEmitente,
            nomeTomador: nomeTomador
        };
    } catch (e) {
        return null; // Silencia erros de estrutura
    }
}

// --- NÚCLEO DE BUSCA ---

async function iniciarCaptura() {
    console.log("====================================================");
    console.log("   SISTEMA DE CAPTURA NFS-E NACIONAL v3.0 - MB");
    console.log("====================================================\n");

    const competenciaReq = readline.question("Digite o mes/ano para AUDITORIA (Ex: 01/2026): ");
    const [mesAlvo, anoAlvo] = competenciaReq.split('/');
    
    let controleNSU = carregarControleNSU();

    for (const cliente of listaClientes) {
        console.log(`\n▶️ Cliente: ${cliente.nome}`);
        
        if (!controleNSU[cliente.cnpj]) controleNSU[cliente.cnpj] = 0;
        let nsuAtual = controleNSU[cliente.cnpj];
        let interromperBusca = false;
        
        let notasDestaSessao = [];
        let mesesExtrasCount = {};

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
                            const extrato = processarNota(nota, cliente);
                            if (extrato) {
                                notasDestaSessao.push(extrato);
                                
                                // Se for de outro mês, adiciona no contador de alertas
                                if (extrato.mesAno !== `${mesAlvo}/${anoAlvo}`) {
                                    mesesExtrasCount[extrato.mesAno] = (mesesExtrasCount[extrato.mesAno] || 0) + 1;
                                }
                            }
                            nsuAtual = nota.NSU; // A esteira anda sempre para frente
                        }
                    } else {
                        interromperBusca = true; 
                    }
                } catch (apiErro) {
                    if (apiErro.response && apiErro.response.status === 404) {
                        interromperBusca = true; 
                    } else { throw apiErro; }
                }
            }

            controleNSU[cliente.cnpj] = nsuAtual;
            salvarControleNSU(controleNSU);

            console.log(`   ✅ Fila do Governo finalizada. Total processado: ${notasDestaSessao.length} nota(s).`);
            
            // Gera o Excel com Resumo e Alertas apenas se encontrou algo
            if (notasDestaSessao.length > 0) {
                await gerarRelatorioExcel(cliente, mesAlvo, anoAlvo, notasDestaSessao, mesesExtrasCount);
                console.log(`   📊 Excel gerado na pasta: ./notas_baixadas/${cliente.nome}/Relatorios/`);
            }

        } catch (erro) {
            console.error(`   ❌ Falha Crítica:`, erro.message);
        }
    }
    console.log("\n====================================================");
    console.log("   PROCESSO FINALIZADO - MB CONTABILIDADE");
    console.log("====================================================");
}

iniciarCaptura();