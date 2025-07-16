/**
 * SCRIPT GERADOR DE RANKING DE VENDAS - NODE.JS
 * SISTEMA DE METAS POR VALOR (R$)
 * 
 * Para usar:
 * 1. npm install
 * 2. Coloque o arquivo EXPORTAR.XLS no mesmo diretório
 * 3. Execute: node ranking.js
 */

const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');

const CONFIG = {
    nomeArquivo: 'EXPORTAR.XLS',
    mesAnalise: '',
    anoAnalise: '',
    
    vendedores: {
        'PrimeiroNome': 'Nome Completo',
        'Fulano': 'Fulano da Silva',
        'Teste': 'Teste de Vendedor'
    },
    
    metas: {
        'PrimeiroNome': 18000.00,  
        'Fulano': 18000.00,  
        'Teste': 18000.00
    }
};

const MESES_OPCOES = {
    1: { codigo: 'Jan', nome: 'Janeiro' },
    2: { codigo: 'Feb', nome: 'Fevereiro' },
    3: { codigo: 'Mar', nome: 'Março' },
    4: { codigo: 'Apr', nome: 'Abril' },
    5: { codigo: 'May', nome: 'Maio' },
    6: { codigo: 'Jun', nome: 'Junho' },
    7: { codigo: 'Jul', nome: 'Julho' },
    8: { codigo: 'Aug', nome: 'Agosto' },
    9: { codigo: 'Sep', nome: 'Setembro' },
    10: { codigo: 'Oct', nome: 'Outubro' },
    11: { codigo: 'Nov', nome: 'Novembro' },
    12: { codigo: 'Dec', nome: 'Dezembro' }
};

function isDataNoMes(dateString, mes, ano) {
    if (!dateString || dateString.includes('-   -')) return false;
    const [dia, mesData, anoData] = dateString.split('-');
    return mesData === mes && anoData === ano;
}

function isVendaCadastradaRelevante(row, mes, ano) {
    const dataLancamento = row.data_lancamento;
    const dataAtivacao = row.data_primeira_ativacao;
    
    if (isDataNoMes(dataLancamento, mes, ano)) return true;
    if (!dataAtivacao || dataAtivacao.includes('-   -')) return true;
    
    return false;
}

function calcularMeta(nomeVendedor, valorInstalado) {
    const metaValor = CONFIG.metas[nomeVendedor] || 18000.00;
    const percentual = Math.min((valorInstalado / metaValor) * 100, 100);
    const percentualReal = (valorInstalado / metaValor) * 100;
    
    return {
        metaValor: metaValor,
        percentual: percentual,
        percentualReal: percentualReal,
        atingiu: valorInstalado >= metaValor,
        valorFaltante: Math.max(metaValor - valorInstalado, 0)
    };
}

function processarDadosVendedor(data, nomeCompleto, mes, ano) {
    const vendasVendedor = data.filter(row => row.vendedor === nomeCompleto);
    
    const nomeVendedor = Object.keys(CONFIG.vendedores).find(
        key => CONFIG.vendedores[key] === nomeCompleto
    ) || nomeCompleto;
    
    const vendasCadastradas = vendasVendedor.filter(row => 
        isVendaCadastradaRelevante(row, mes, ano)
    );
    
    const vendasInstaladaMes = vendasVendedor.filter(row => 
        isDataNoMes(row.data_primeira_ativacao, mes, ano)
    );
    
    const valorInstaladoMes = vendasInstaladaMes.reduce((sum, row) => 
        sum + (parseFloat(row.valor_final) || 0), 0
    );
    
    const dadosMeta = calcularMeta(nomeVendedor, valorInstaladoMes);
    
    return {
        nomeCompleto,
        vendasCadastradas: vendasCadastradas.length,
        vendasInstaladaMes: vendasInstaladaMes.length,
        valorInstaladoMes: valorInstaladoMes,
        meta: dadosMeta,
        detalhesVendasInstaladas: vendasInstaladaMes.map(venda => ({
            cliente: venda.nome_cliente,
            valor: parseFloat(venda.valor_final) || 0,
            dataAtivacao: venda.data_primeira_ativacao,
            plano: venda.nome_do_plano
        }))
    };
}

function gerarTextoRanking(ranking, mes, ano) {
    const mesAnoCompleto = `${mes}/20${ano}`;
    let texto = `🏆 RANKING DE DESEMPENHO DE VENDAS - ${mesAnoCompleto.toUpperCase()}\n`;
    texto += `Organizado por maior valor de vendas instaladas em ${mesAnoCompleto}\n\n`;
    
    ranking.forEach((vendedor, index) => {
        const posicao = index + 1;
        const medalha = posicao === 1 ? '🥇' : posicao === 2 ? '🥈' : posicao === 3 ? '🥉' : '📍';
        
        texto += `${medalha} ${posicao}º LUGAR - ${vendedor.nome.toUpperCase()}\n`;
        texto += `   Vendas Instaladas em ${mesAnoCompleto}: ${vendedor.vendasInstaladaMes}\n`;
        texto += `   Valor das Vendas Instaladas: R$ ${vendedor.valorInstaladoMes.toFixed(2).replace('.', ',')}\n`;
        
        if (vendedor.meta) {
            texto += `   Meta de Valor: R$ ${vendedor.meta.metaValor.toFixed(2).replace('.', ',')} (${vendedor.meta.percentualReal.toFixed(1)}%)\n`;
            if (vendedor.meta.atingiu) {
                texto += `   ✅ META ATINGIDA! ${vendedor.meta.percentualReal > 100 ? 'SUPEROU a meta!' : ''}\n`;
            } else {
                texto += `   ⏳ Faltam R$ ${vendedor.meta.valorFaltante.toFixed(2).replace('.', ',')} para atingir a meta\n`;
            }
        }
        texto += `\n`;
    });
    
    return texto;
}

function mesParaPortugues(mes) {
    const meses = {
        'Jan': 'Janeiro', 'Feb': 'Fevereiro', 'Mar': 'Março',
        'Apr': 'Abril', 'May': 'Maio', 'Jun': 'Junho',
        'Jul': 'Julho', 'Aug': 'Agosto', 'Sep': 'Setembro',
        'Oct': 'Outubro', 'Nov': 'Novembro', 'Dec': 'Dezembro'
    };
    return meses[mes] || mes;
}

function gerarHTMLRanking(ranking, mes, ano) {
    const mesPortugues = mesParaPortugues(mes);
    const anoCompleto = `20${ano}`;
    
    let html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ranking de Vendas - ${mesPortugues} ${anoCompleto}</title>
    <style>
        * { 
            margin: 0; 
            padding: 0; 
            box-sizing: border-box; 
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #3b73c7 100%);
            min-height: 100vh;
            padding: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        
        .container {
            background: linear-gradient(135deg, rgba(74, 111, 165, 0.9) 0%, rgba(59, 89, 152, 0.9) 100%);
            border-radius: 20px;
            padding: 30px;
            max-width: 650px;
            width: 100%;
            box-shadow: 0 25px 80px rgba(0, 0, 0, 0.4);
            border: 2px solid rgba(255, 255, 255, 0.1);
            backdrop-filter: blur(20px);
        }
        
        .header {
            text-align: center;
            margin-bottom: 25px;
            color: white;
        }
        
        .logo {
            max-width: 250px;
            margin-bottom: 15px;
        }
        
        .title {
            font-size: 24px;
            font-weight: 700;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
        }
        
        .subtitle {
            font-size: 16px;
            font-weight: 400;
            opacity: 0.9;
            text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
        }
        
        .ranking {
            display: flex;
            flex-direction: column;
            gap: 14px;
        }
        
        .vendedor-card {
            position: relative;
            border-radius: 20px;
            padding: 18px;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            backdrop-filter: blur(20px);
            border: 1px solid rgba(255, 255, 255, 0.3);
            box-shadow: 0 6px 24px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }
        
        .vendedor-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(
                90deg, 
                rgba(46, 204, 113, 0.3) 0%, 
                rgba(46, 204, 113, 0.2) var(--progress), 
                rgba(255, 255, 255, 0.15) var(--progress), 
                rgba(255, 255, 255, 0.15) 100%
            );
            z-index: 1;
            border-radius: 20px;
        }
        
        .vendedor-card.meta-baixa::before {
            background: linear-gradient(
                90deg, 
                rgba(110, 95, 245, 0.2) 0%, 
                rgba(110, 95, 245, 0.2) var(--progress), 
                rgba(255, 255, 255, 0.15) var(--progress), 
                rgba(255, 255, 255, 0.15) 100%
            );
        }
        
        .vendedor-card.meta-media::before {
            background: linear-gradient(
                90deg, 
                rgba(34, 217, 230, 0.3) 0%, 
                rgba(34, 217, 230, 0.3) var(--progress), 
                rgba(255, 255, 255, 0.15) var(--progress), 
                rgba(255, 255, 255, 0.15) 100%
            );
        }
        
        .vendedor-card.meta-alta::before {
            background: linear-gradient(
                90deg, 
                rgba(155, 89, 182, 0.3) 0%, 
                rgba(155, 89, 182, 0.2) var(--progress), 
                rgba(255, 255, 255, 0.15) var(--progress), 
                rgba(255, 255, 255, 0.15) 100%
            );
        }
        
        .vendedor-card.meta-superou::before {
            background: linear-gradient(
                90deg, 
                rgba(46, 204, 113, 0.4) 0%, 
                rgba(46, 204, 113, 0.3) 100%
            );
        }
        
        .vendedor-card > * {
            position: relative;
            z-index: 2;
        }
        
        .vendedor-card:hover {
            transform: translateY(-8px) scale(1.02);
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.25);
            border: 1px solid rgba(255, 255, 255, 0.4);
        }
        
        .vendedor-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
        }
        
        .vendedor-info {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .posicao {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 35px;
            height: 35px;
            border-radius: 50%;
            color: white;
            font-weight: 700;
            font-size: 16px;
            text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
        }
        
        .emoji-badge {
            font-size: 20px;
            margin-right: 6px;
        }
        
        .nome {
            font-size: 25px;
            font-weight: 600;
            color: rgba(255, 255, 255, 0.95);
            text-transform: capitalize;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
        }
        
        .valor-principal {
            font-size: 20px;
            font-weight: 700;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
            padding: 6px 14px;
            border-radius: 10px;
            backdrop-filter: blur(10px);
        }
        
        .metricas-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-top: 12px;
        }
        
        .metrica-box {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 14px;
            padding: 14px;
            text-align: center;
            border: 1px solid rgba(255, 255, 255, 0.2);
            backdrop-filter: blur(10px);
            transition: all 0.3s ease;
        }
        
        .metrica-box:hover {
            background: rgba(255, 255, 255, 0.2);
            transform: translateY(-2px);
        }
        
        .metrica-label {
            font-size: 11px;
            color: rgba(255, 255, 255, 0.8);
            font-weight: 500;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .metrica-valor {
            font-size: 22px;
            font-weight: 700;
            color: rgba(255, 255, 255, 0.95);
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
        }
        
        .metrica-meta .metrica-valor {
            font-size: 22px;
        }
        
        .primeiro .posicao { background: linear-gradient(135deg, #f1c40f, #f39c12); }
        .segundo .posicao { background: linear-gradient(135deg, #95a5a6, #7f8c8d); }
        .terceiro .posicao { background: linear-gradient(135deg, #cd7f32, #b8651b); }
        .quarto .posicao { background: linear-gradient(135deg, #3498db, #2980b9); }
        .quinto .posicao { background: linear-gradient(135deg, #3498db, #2980b9); }
        
        .primeiro .valor-principal { 
            color: #ffd700; 
            background: rgba(255, 215, 0, 0.15); 
            border: 1px solid rgba(255, 215, 0, 0.3); 
        }
        .segundo .valor-principal { 
            color:rgb(209, 209, 209); 
            background: rgba(160, 160, 160, 0.15); 
            border: 1px solid rgba(192, 192, 192, 0.3); 
        }
        .terceiro .valor-principal { 
            color:rgb(224, 154, 84); 
            background: rgba(205, 127, 50, 0.15); 
            border: 1px solid rgba(205, 127, 50, 0.3); 
        }
        .quarto .valor-principal, .quinto .valor-principal { 
            color: rgba(255, 255, 255, 0.95); 
            background: rgba(255, 255, 255, 0.1); 
            border: 1px solid rgba(255, 255, 255, 0.2); 
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                padding: 20px;
            }
            
            .vendedor-header {
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }
            
            .valor-principal {
                font-size: 18px;
            }
            
            .metricas-grid {
                grid-template-columns: 1fr;
                gap: 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <img src="assets/logo-white.png" class="logo" alt="Logo">
            <h1 class="title">Ranking de Vendas</h1>
            <p class="subtitle">${mesPortugues} ${anoCompleto} - Desempenho por Vendedor</p>
        </div>
        
        <div class="ranking">`;

    ranking.forEach((vendedor, index) => {
        const posicao = index + 1;
        const classes = ['primeiro', 'segundo', 'terceiro', 'quarto', 'quinto'];
        const emojis = ['🥇', '🥈', '🥉'];
        const classe = classes[index] || 'outros';
        const emoji = emojis[index] || "     ";
        
        const dadosMeta = vendedor.meta;
        
        let classeMeta = '';
        if (dadosMeta.percentualReal >= 100) {
            classeMeta = 'meta-superou';
        } else if (dadosMeta.percentual >= 80) {
            classeMeta = 'meta-alta';
        } else if (dadosMeta.percentual >= 50) {
            classeMeta = 'meta-media';
        } else {
            classeMeta = 'meta-baixa';
        }
        
        html += `
            <div class="vendedor-card ${classe} ${classeMeta}" style="--progress: ${dadosMeta.percentual}%">
                <div class="vendedor-header">
                    <div class="vendedor-info">
                        <span class="posicao">${posicao}°</span>
                        <span class="emoji-badge">${emoji}</span>
                        <span class="nome">${vendedor.nome}</span>
                    </div>
                    <div class="valor-principal">R$ ${vendedor.valorInstaladoMes.toFixed(2).replace('.', ',')}</div>
                </div>
                
                <div class="metricas-grid">
                    <div class="metrica-box">
                        <div class="metrica-label">Instaladas em ${mesPortugues}</div>
                        <div class="metrica-valor">${vendedor.vendasInstaladaMes}</div>
                    </div>
                    <div class="metrica-box metrica-meta">
                        <div class="metrica-label">Meta | R$${(dadosMeta.metaValor/1000).toFixed(0)}k</div>
                        <div class="metrica-valor">${dadosMeta.percentualReal.toFixed(0)}%</div>
                    </div>
                </div>
            </div>`;
    });

    html += `
        </div>
    </div>
</body>
</html>`;

    return html;
}

function lerArquivo(nomeArquivo) {
    try {
        if (!fs.existsSync(nomeArquivo)) {
            throw new Error(`Arquivo "${nomeArquivo}" não encontrado no diretório atual.`);
        }
        
        const fileBuffer = fs.readFileSync(nomeArquivo);
        return XLSX.read(fileBuffer, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true
        });
    } catch (error) {
        if (error.code === 'MODULE_NOT_FOUND') {
            throw new Error('❌ Erro: Instale as dependências com: npm install xlsx');
        }
        throw error;
    }
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Função para criar interface CLI
function criarInterface() {
    return readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
}

// Função para fazer pergunta com prompt customizado
function pergunta(rl, texto) {
    return new Promise((resolve) => {
        rl.question(texto, (resposta) => {
            resolve(resposta.trim());
        });
    });
}

// CLI para seleção de mês e ano
async function selecionarPeriodo() {
    const rl = criarInterface();
    
    console.log('\n🗓️  SELEÇÃO DE PERÍODO PARA ANÁLISE');
    console.log('═'.repeat(45));
    
    // Mostrar opções de meses
    console.log('\n📅 MESES DISPONÍVEIS:');
    Object.entries(MESES_OPCOES).forEach(([num, mes]) => {
        console.log(`   ${num.padStart(2)} - ${mes.nome}`);
    });
    
    let mesEscolhido, anoEscolhido;
    
    // Loop para validar mês
    while (true) {
        const respostaMes = await pergunta(rl, '\n🔸 Digite o número do mês (1-12): ');
        const numeroMes = parseInt(respostaMes);
        
        if (numeroMes >= 1 && numeroMes <= 12) {
            mesEscolhido = MESES_OPCOES[numeroMes];
            console.log(`✅ Mês selecionado: ${mesEscolhido.nome}`);
            break;
        } else {
            console.log('❌ Mês inválido! Digite um número entre 1 e 12.');
        }
    }
    
    // Loop para validar ano
    while (true) {
        const respostaAno = await pergunta(rl, '\n🔸 Digite o ano (ex: 2025 ou 25): ');
        
        let anoFormatado;
        
        // Aceitar formato completo (2025) ou abreviado (25)
        if (respostaAno.length === 4 && !isNaN(respostaAno)) {
            // Formato completo: 2025 -> 25
            const anoCompleto = parseInt(respostaAno);
            if (anoCompleto >= 2000 && anoCompleto <= 2099) {
                anoFormatado = respostaAno.slice(2); // Pega os últimos 2 dígitos
            } else {
                console.log('❌ Ano inválido! Digite um ano entre 2000 e 2099.');
                continue;
            }
        } else if (respostaAno.length === 2 && !isNaN(respostaAno)) {
            // Formato abreviado: 25
            const anoAbrev = parseInt(respostaAno);
            if (anoAbrev >= 0 && anoAbrev <= 99) {
                anoFormatado = respostaAno.padStart(2, '0'); // Garante 2 dígitos
            } else {
                console.log('❌ Ano inválido! Digite no formato YY (ex: 25) ou YYYY (ex: 2025).');
                continue;
            }
        } else {
            console.log('❌ Formato inválido! Digite no formato YY (ex: 25) ou YYYY (ex: 2025).');
            continue;
        }
        
        anoEscolhido = anoFormatado;
        const anoExibicao = `20${anoFormatado}`;
        console.log(`✅ Ano selecionado: ${anoExibicao}`);
        break;
    }
    
    // Opção para arquivo personalizado
    console.log('\n📁 ARQUIVO DE DADOS:');
    const respostaArquivo = await pergunta(rl, `🔸 Pressione ENTER para usar "${CONFIG.nomeArquivo}" ou digite outro nome: `);
    
    const arquivoEscolhido = respostaArquivo || CONFIG.nomeArquivo;
    if (respostaArquivo) {
        console.log(`✅ Arquivo personalizado: ${arquivoEscolhido}`);
    } else {
        console.log(`✅ Usando arquivo padrão: ${arquivoEscolhido}`);
    }
    
    rl.close();
    
    // Atualizar configurações
    CONFIG.mesAnalise = mesEscolhido.codigo;
    CONFIG.anoAnalise = anoEscolhido;
    CONFIG.nomeArquivo = arquivoEscolhido;
    
    console.log('\n📋 RESUMO DA CONFIGURAÇÃO:');
    console.log(`   📅 Período: ${mesEscolhido.nome}/20${anoEscolhido}`);
    console.log(`   📁 Arquivo: ${arquivoEscolhido}`);
    console.log(`   👥 Vendedores: ${Object.keys(CONFIG.vendedores).length} configurados`);
    
    await delay(1000);
    
    return {
        mes: mesEscolhido,
        ano: anoEscolhido,
        arquivo: arquivoEscolhido
    };
}

async function animacaoCarregamento(texto, duracao = 2000) {
    const frames = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏'];
    const interval = 80;
    const cycles = Math.floor(duracao / interval);
    
    for (let i = 0; i < cycles; i++) {
        const frame = frames[i % frames.length];
        process.stdout.write(`\r${frame} ${texto}`);
        await delay(interval);
    }
    process.stdout.write('\r' + ' '.repeat(texto.length + 2) + '\r');
}

function mostrarProgresso(progresso, texto = '') {
    const largura = 30;
    const preenchido = Math.floor((progresso / 100) * largura);
    const vazio = largura - preenchido;
    
    const barra = '█'.repeat(preenchido) + '░'.repeat(vazio);
    const porcentagem = progresso.toFixed(0).padStart(3);
    
    process.stdout.write(`\r📊 [${barra}] ${porcentagem}% ${texto}`);
    
    if (progresso >= 100) {
        console.log('');
    }
}

async function gerarRankingVendas(arquivo = null) {
    try {
        console.log('🚀 INICIANDO ANÁLISE DE VENDAS');
        console.log('═'.repeat(50));
        await delay(500);
        
        await animacaoCarregamento('Preparando ambiente...', 1500);
        console.log('✅ Ambiente preparado!');
        await delay(300);
        
        const nomeArquivo = arquivo || CONFIG.nomeArquivo;
        await animacaoCarregamento(`Localizando arquivo: ${nomeArquivo}`, 1000);
        console.log(`📂 Arquivo encontrado: ${nomeArquivo}`);
        await delay(300);
        
        await animacaoCarregamento('Carregando dados do Excel...', 1500);
        const workbook = lerArquivo(nomeArquivo);
        
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(firstSheet);
        
        console.log(`📊 Dados carregados com sucesso: ${data.length} registros`);
        await delay(500);
        
        const rankingData = {};
        const vendedoresList = Object.keys(CONFIG.vendedores);
        
        console.log(`\n👥 Processando ${vendedoresList.length} vendedores:`);
        console.log('─'.repeat(40));
        await delay(400);
        
        for (let i = 0; i < vendedoresList.length; i++) {
            const vendedor = vendedoresList[i];
            const nomeCompleto = CONFIG.vendedores[vendedor];
            
            mostrarProgresso(((i) / vendedoresList.length) * 100, `Processando ${vendedor}...`);
            await delay(300);
            
            await animacaoCarregamento(`Analisando vendas de ${vendedor}`, 800);
            
            const dadosProcessados = processarDadosVendedor(
                data, 
                nomeCompleto, 
                CONFIG.mesAnalise, 
                CONFIG.anoAnalise
            );
            
            rankingData[vendedor] = dadosProcessados;
            
            mostrarProgresso(((i + 1) / vendedoresList.length) * 100, 'Concluído!');
            
            console.log(`✅ ${vendedor}: ${dadosProcessados.vendasInstaladaMes} vendas instaladas, R$ ${dadosProcessados.valorInstaladoMes.toFixed(2)} (Meta: ${dadosProcessados.meta.percentualReal.toFixed(1)}%)`);
            await delay(400);
        }
        
        console.log('\n🔄 Organizando ranking...');
        await animacaoCarregamento('Calculando posições...', 1200);
        
        const rankingArray = Object.entries(rankingData)
            .map(([nome, dados]) => ({ nome, ...dados }))
            .sort((a, b) => b.valorInstaladoMes - a.valorInstaladoMes);
        
        console.log('✅ Ranking calculado!');
        await delay(300);
        
        console.log('\n📄 Gerando relatórios...');
        await animacaoCarregamento('Criando arquivo de texto...', 800);
        const textoRanking = gerarTextoRanking(rankingArray, CONFIG.mesAnalise, CONFIG.anoAnalise);
        
        await animacaoCarregamento('Criando arquivo HTML...', 1000);
        const htmlRanking = gerarHTMLRanking(rankingArray, CONFIG.mesAnalise, CONFIG.anoAnalise);
        
        console.log('✅ Relatórios gerados!');
        await delay(300);
        
        console.log('\n🏆 RESULTADO FINAL:');
        console.log('═'.repeat(50));
        await delay(800);
        
        console.log(textoRanking);
        
        console.log('\n💾 Salvando arquivos...');
        await animacaoCarregamento('Escrevendo arquivos no disco...', 1000);
        
        const nomeArquivoTxt = `ranking_${CONFIG.mesAnalise}_${CONFIG.anoAnalise}.txt`;
        const nomeArquivoHtml = `ranking_${CONFIG.mesAnalise}_${CONFIG.anoAnalise}.html`;
        
        fs.writeFileSync(nomeArquivoTxt, textoRanking);
        fs.writeFileSync(nomeArquivoHtml, htmlRanking);
        
        console.log(`✅ Arquivos salvos com sucesso!`);
        console.log(`   📄 ${nomeArquivoTxt}`);
        console.log(`   🌐 ${nomeArquivoHtml}`);
        
        console.log('\n🎉 ANÁLISE CONCLUÍDA COM SUCESSO!');
        console.log('═'.repeat(50));
        
        global.rankingResultados = {
            dados: rankingArray,
            texto: textoRanking,
            html: htmlRanking,
            config: CONFIG
        };
        
        return rankingArray;
        
    } catch (error) {
        console.log('\n❌ ERRO DURANTE A ANÁLISE');
        console.log('═'.repeat(50));
        console.error(`Detalhes: ${error.message}`);
        throw error;
    }
}

function verDetalhesVendedor(nomeVendedor) {
    if (!global.rankingResultados) {
        console.error('❌ Execute gerarRankingVendas() primeiro!');
        return;
    }
    
    const vendedor = global.rankingResultados.dados.find(v => v.nome === nomeVendedor);
    if (!vendedor) {
        console.error(`❌ Vendedor "${nomeVendedor}" não encontrado!`);
        return;
    }
    
    console.log(`\n📋 DETALHES - ${vendedor.nome.toUpperCase()}`);
    console.log(`├─ Vendas Instaladas: ${vendedor.vendasInstaladaMes}`);
    console.log(`├─ Valor Total: R$ ${vendedor.valorInstaladoMes.toFixed(2)}`);
    console.log(`├─ Meta de Valor: R$ ${vendedor.meta.metaValor.toFixed(2)}`);
    console.log(`├─ Progresso: ${vendedor.meta.percentualReal.toFixed(1)}%`);
    console.log(`└─ Status: ${vendedor.meta.atingiu ? '✅ META ATINGIDA!' : `⏳ Faltam R$ ${vendedor.meta.valorFaltante.toFixed(2)}`}`);
    
    if (vendedor.detalhesVendasInstaladas.length > 0) {
        console.log('\n💰 VENDAS INSTALADAS:');
        vendedor.detalhesVendasInstaladas.forEach((venda, idx) => {
            console.log(`${idx + 1}. ${venda.cliente} - R$ ${venda.valor.toFixed(2)} (${venda.dataAtivacao})`);
        });
    }
}

function verResumoMetas() {
    if (!global.rankingResultados) {
        console.error('❌ Execute gerarRankingVendas() primeiro!');
        return;
    }
    
    console.log('\n🎯 RESUMO DAS METAS POR VALOR');
    console.log('═'.repeat(70));
    
    global.rankingResultados.dados.forEach((vendedor, index) => {
        const status = vendedor.meta.atingiu ? 
            (vendedor.meta.percentualReal > 100 ? '🟣 SUPEROU' : '🟢 ATINGIU') : 
            '🔴 PENDENTE';
        
        const valorFormatado = `R$ ${(vendedor.valorInstaladoMes/1000).toFixed(1)}k`;
        const metaFormatada = `R$ ${(vendedor.meta.metaValor/1000).toFixed(1)}k`;
        
        console.log(`${vendedor.nome.padEnd(10)} | ${valorFormatado.padStart(8)}/${metaFormatada.padStart(8)} | ${vendedor.meta.percentualReal.toFixed(1).padStart(5)}% | ${status}`);
    });
    
    const totalVendedores = global.rankingResultados.dados.length;
    const atingiramMeta = global.rankingResultados.dados.filter(v => v.meta.atingiu).length;
    const superaramMeta = global.rankingResultados.dados.filter(v => v.meta.percentualReal > 100).length;
    
    console.log('─'.repeat(70));
    console.log(`📊 ESTATÍSTICAS:`);
    console.log(`   Atingiram a meta: ${atingiramMeta}/${totalVendedores} (${((atingiramMeta/totalVendedores)*100).toFixed(1)}%)`);
    console.log(`   Superaram a meta: ${superaramMeta}/${totalVendedores} (${((superaramMeta/totalVendedores)*100).toFixed(1)}%)`);
}

console.log(`
🚀 SCRIPT GERADOR DE RANKING DE VENDAS - NODE.JS
📊 SISTEMA DE METAS POR VALOR (R$)
─────────────────────────────────────────────────

📋 DEPENDÊNCIAS:
• npm install xlsx readline-sync

🔧 FUNÇÕES DISPONÍVEIS:
• gerarRankingVendas() - Executa a análise completa
• gerarRankingVendas('ARQUIVO.XLS') - Usa arquivo específico  
• verDetalhesVendedor('NomeVendedor') - Mostra detalhes de um vendedor
• verResumoMetas() - Mostra resumo das metas de todos

⚙️ VENDEDORES CONFIGURADOS:
• ${Object.keys(CONFIG.vendedores).join(', ')}
• Metas por Valor: ${Object.entries(CONFIG.metas).map(([nome, meta]) => `${nome} (R$ ${(meta/1000).toFixed(1)}k)`).join(', ')}

🔄 Para alterar configurações, edite o objeto CONFIG no início do script.
`);

(async () => {
    try {
        // Executar CLI para seleção de período
        await selecionarPeriodo();
        
        // Executar análise com as configurações selecionadas
        await gerarRankingVendas();
    } catch (error) {
        console.error('\n❌ Erro durante a execução:', error.message);
        process.exit(1);
    }
})();

module.exports = {
    gerarRankingVendas,
    verDetalhesVendedor,
    verResumoMetas,
    CONFIG
};