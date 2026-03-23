// URLs das suas Planilhas
const URL_COMISSOES = "https://script.google.com/macros/s/AKfycbxZO7jiev9MWxuQgAfaPdark6geTUqWAr9eYypWG_0qvqx-sVxdx6agnDu1aFIMwBL6aA/exec";
const URL_FORNECEDORES = "https://script.google.com/macros/s/AKfycbxpMXA3xWANJ8ivdoj_3ZbUV0nCXDWvJ7Ja5E6bTAdVquSImH_gDfQ9pabnwvoaZK5b/exec";
const URL_SHEET_BANCO = "https://docs.google.com/spreadsheets/d/1_UIvezU3eh5HQ98ttIXsViCCsY2opGwNOfZbv4SVFfc/edit?usp=sharing";
const URL_SHEET_LOGISTICA = "https://docs.google.com/spreadsheets/d/1inVjNncz3YdWV31iEShiYjCUkWEE0fOfkTXCwRDu98k/edit?usp=sharing";
const URL_SHEET_GERENCIAL = "https://docs.google.com/spreadsheets/d/17PYbOV8CuEwghbaDiUmJXvc1mCT7tZ55iKvkQFzTeXc/edit?usp=sharing";
const URL_LOGIN_DB = "https://script.google.com/macros/s/AKfycbyffqQQUSRWVVpyQyKyKTC5fwyEii8RzF9fFlJflwhFupAZ-QusTzhXrGSgMFEZQRHgxA/exec";

// 👇 COLOQUE AQUI O ID DA SUA PLANILHA ONDE AS COMISSÕES SÃO SALVAS 👇
const ID_PLANILHA_COMISSOES = "17PYbOV8CuEwghbaDiUmJXvc1mCT7tZ55iKvkQFzTeXc";

let usuarioLogado = localStorage.getItem('usuarioLogado');

let dadosEmpresa = {}; 
let dadosExtras = {};
let currentSheetUrl = "";
let totalVendidoSessao = 0;
let totalComissaoSessao = 0;
let chartMensal = null;
let chartVendedoras = null;

// Função para criar o Hash da Senha (SHA-256)
async function hashPassword(message) {
    const msgUint8 = new TextEncoder().encode(message);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgUint8);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

function setAndLockVendedora(user) {
    const vendedoraSelect = document.getElementById('vendedora');
    if (!vendedoraSelect) return;

    const admins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE', 'RENATA', 'CAROL'];
    const isAdmin = admins.some(adm => adm.toUpperCase() === user.toUpperCase());
    let optionExists = Array.from(vendedoraSelect.options).some(opt => opt.value === user);

    if (!optionExists) {
        const newOption = new Option(user, user);
        vendedoraSelect.add(newOption);
    }

    vendedoraSelect.value = user;
    vendedoraSelect.disabled = !isAdmin;

    // Aciona a busca do Dashboard real após identificar a vendedora
    setTimeout(carregarDashboardReal, 1000); 
    if(!window.dashInterval) {
        window.dashInterval = setInterval(carregarDashboardReal, 60000); // Atualiza os gráficos a cada 1 min
    }
}

async function realizarLogin() {
    const user = document.getElementById('userLogin').value.trim();
    const senhaRaw = document.getElementById('senhaLogin').value;
    
    const senhaHash = await hashPassword(senhaRaw);
    const status = document.getElementById('loginStatus');

    status.innerText = "⏳ Verificando na base de dados...";
    status.style.display = 'block';
    status.style.color = 'blue';

    try {
        const response = await fetch(URL_LOGIN_DB, {
            method: 'POST',
            body: JSON.stringify({ acao: "login", nome: user, senha: senhaHash })
        });
        const res = await response.text();

        if (res === "Autorizado") {
            document.getElementById('telaLogin').style.display = 'none';
            localStorage.setItem('usuarioLogado', user);
            usuarioLogado = user;
            document.getElementById('nomeUsuarioHeader').innerHTML = `<i class="fas fa-user-circle"></i> ${user}`;
            setAndLockVendedora(user);
        } else {
            status.innerText = "❌ Usuário ou senha incorretos!";
            status.style.color = 'red';
        }
    } catch (e) {
        status.innerText = "❌ Erro de conexão!";
        status.style.color = 'red';
    }
}

if (usuarioLogado) {
    document.getElementById('telaLogin').style.display = 'none';
    document.getElementById('nomeUsuarioHeader').innerHTML = `<i class="fas fa-user-circle"></i> ${usuarioLogado}`;
    setTimeout(() => setAndLockVendedora(usuarioLogado), 500);
}

async function cadastrarNovoVendedor() {
    const nome = document.getElementById('novoNome').value.trim();
    const senhaRaw = document.getElementById('novaSenha').value;

    if (!nome || !senhaRaw) return Swal.fire('Atenção', 'Por favor, preencha nome e senha.', 'warning');
    
    const senhaHash = await hashPassword(senhaRaw);

    try {
        const response = await fetch(URL_LOGIN_DB, {
            method: 'POST',
            body: JSON.stringify({ acao: "cadastrar", nome: nome, senha: senhaHash })
        });
        const texto = await response.text();
        
        if (texto === "Sucesso") {
            Swal.fire('Sucesso!', "✅ USUARIO " + nome + " cadastrado com sucesso!", 'success');
            alternarAbaAuth('entrar');
        }
    } catch (e) {
        Swal.fire('Erro', 'Erro ao conectar com a planilha de acesso.', 'error');
    }
}

async function efetuarLogin() {
    const user = document.getElementById('selectUser').value;
    const senhaRaw = document.getElementById('senhaInput').value;
    
    const senhaHash = await hashPassword(senhaRaw);

    try {
        const response = await fetch(URL_LOGIN_DB, {
            method: 'POST',
            body: JSON.stringify({ acao: "login", nome: user, senha: senhaHash })
        });
        const resultado = await response.text();

        if (resultado === "Autorizado") {
            salvarSessao(user);
        } else {
            Swal.fire('Erro', "❌ Senha incorreta para " + user, 'error');
        }
    } catch (e) {
        Swal.fire('Erro', "Erro ao validar login no banco de dados.", 'error');
    }
}

function salvarSessao(user) {
    localStorage.setItem('usuarioLogado', user);
    location.reload(); 
}

function abrirModalLogin() {
    document.getElementById('modalLogin').classList.remove('hidden');
    
    if (usuarioLogado) {
        document.getElementById('camposLogin').classList.add('hidden');
        document.getElementById('areaLogado').classList.remove('hidden');
        document.getElementById('nomeLogado').innerText = usuarioLogado;
        document.getElementById('tituloLogin').innerText = "Meu Perfil";
    } else {
        document.getElementById('camposLogin').classList.remove('hidden');
        document.getElementById('areaLogado').classList.add('hidden');
        document.getElementById('tituloLogin').innerText = "Acesso Miss Rôse";
    }
}

function fecharModalLogin() {
    document.getElementById('modalLogin').classList.add('hidden');
}

function mostrarSenha() {
    const s = document.getElementById('senhaInput');
    const i = document.getElementById('toggleSenha');
    if (s.type === "password") {
        s.type = "text";
        i.classList.replace('fa-eye', 'fa-eye-slash');
    } else {
        s.type = "password";
        i.classList.replace('fa-eye-slash', 'fa-eye');
    }
}

function alternarAbaAuth(tipo) {
    const btnE = document.getElementById('btnTabEntrar');
    const btnC = document.getElementById('btnTabCadastrar');
    const formE = document.getElementById('formEntrar');
    const formC = document.getElementById('formCadastrar');
    
    if(tipo === 'entrar') {
        btnE.style.background = 'var(--primary)'; btnE.style.color = 'white';
        btnC.style.background = '#ccc'; btnC.style.color = '#333';
        formE.classList.remove('hidden'); formC.classList.add('hidden');
        document.getElementById('tituloAuth').innerText = "Acesso Restrito";
    } else {
        btnC.style.background = 'var(--primary)'; btnC.style.color = 'white';
        btnE.style.background = '#ccc'; btnE.style.color = '#333';
        formC.classList.remove('hidden'); formE.classList.add('hidden');
        document.getElementById('tituloAuth').innerText = "Novo Vendedor";
    }
}

function efetuarLogout() {
    localStorage.removeItem('usuarioLogado');
    location.reload(); 
}

function alternarTela(tela) {
    const telaDash = document.getElementById('telaDashboard');
    const telaCom = document.getElementById('telaComissoes');
    const telaFor = document.getElementById('telaFornecedores');
    const telaPlan = document.getElementById('telaPlanilhas');
    const navDash = document.getElementById('navDashboard');
    const navCom = document.getElementById('navComissoes');
    const navFor = document.getElementById('navFornecedores');
    const navPlan = document.getElementById('navPlanilhas');

    if (telaDash) telaDash.classList.add('hidden');
    telaCom.classList.add('hidden');
    telaFor.classList.add('hidden');
    telaPlan.classList.add('hidden');
    
    if (navDash) navDash.classList.remove('active');
    navCom.classList.remove('active');
    navFor.classList.remove('active');
    navPlan.classList.remove('active');

    if (tela === 'dashboard') {
        if (telaDash) telaDash.classList.remove('hidden');
        if (navDash) navDash.classList.add('active');
        
        // Dá um tempo para o navegador pintar a tela antes de expandir os gráficos
        setTimeout(() => {
            if (chartMensal) { chartMensal.resize(); chartMensal.update(); }
            if (chartVendedoras) { chartVendedoras.resize(); chartVendedoras.update(); }
        }, 150);
    } else if (tela === 'comissoes') {
        telaCom.classList.remove('hidden'); navCom.classList.add('active');
    } else if (tela === 'fornecedores') {
        telaFor.classList.remove('hidden'); navFor.classList.add('active');
    } else if (tela === 'planilhas') {
        telaPlan.classList.remove('hidden'); navPlan.classList.add('active');
    }
}

function abrirPlanilha(tipo) {
    if (tipo === 'gerencial') {
        const admins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE', 'RENATA', 'CAROL'];
        const isAdmin = admins.some(adm => adm.toUpperCase() === (usuarioLogado || "").toUpperCase());
        
        if (!isAdmin) {
            Swal.fire('Acesso Negado', '🚫 Apenas administradores podem acessar a Planilha Gerencial.', 'error');
            return;
        }
        currentSheetUrl = URL_SHEET_GERENCIAL;
    } else if (tipo === 'banco') {
        currentSheetUrl = URL_SHEET_BANCO;
    } else if (tipo === 'logistica') {
        currentSheetUrl = URL_SHEET_LOGISTICA;
    }
    
    document.getElementById('planilhasCards').style.display = 'none';
    document.getElementById('planilhaViewer').style.display = 'flex';
    document.getElementById('framePlanilha').src = currentSheetUrl;
}

function fecharPlanilha() {
    document.getElementById('planilhasCards').style.display = 'grid';
    document.getElementById('planilhaViewer').style.display = 'none';
    document.getElementById('framePlanilha').src = "about:blank";
}

function abrirNovaAba() {
    if(currentSheetUrl) window.open(currentSheetUrl, '_blank');
}

function exibirStatus(id, texto, bg, cor) {
    const div = document.getElementById(id);
    div.innerText = texto; div.style.background = bg; div.style.color = cor; div.style.display = 'block';
}

function toggleRepresentante() {
    const tipo = document.getElementById('tipoVenda').value;
    document.getElementById('campoRepresentante').classList.toggle('hidden', tipo !== 'representante');
}

function mudarCorEquipe() {
    const e = document.getElementById('equipe').value;
    const h = document.getElementById('mainHeader');
    const b = document.getElementById('btnRegistrar');
    const c = document.getElementById('telaComissoes');
    
    const divTipo = document.getElementById('divTipoVendedora');
    if (divTipo) {
        if (e === 'Renata') divTipo.classList.add('hidden');
        else divTipo.classList.remove('hidden');
    }

    const cor = (e === 'Carol') ? '#8e44ad' : '#d81b60';
    const grad = (e === 'Carol') ? 'linear-gradient(135deg, #8e44ad, #6c3483)' : 'linear-gradient(135deg, #d81b60, #ad1457)';
    h.style.background = grad; b.style.backgroundColor = cor; c.style.borderTopColor = cor;
}

document.querySelectorAll('#valorNota, #valorFora, #porcVend, #porcRep').forEach(el => {
    el.addEventListener('input', calcularComissaoRealTime);
});

function calcularComissaoRealTime() {
    const valorNota = parseFloat(document.getElementById('valorNota').value) || 0;
    const valorFora = parseFloat(document.getElementById('valorFora').value) || 0;
    const porcVend = parseFloat(document.getElementById('porcVend').value) || 0;
    const porcRep = parseFloat(document.getElementById('porcRep').value) || 0;
    const tipoVenda = document.getElementById('tipoVenda').value;

    const totalBase = valorNota + valorFora;
    const comissaoVend = (totalBase * porcVend) / 100;
    const comissaoRep = (totalBase * porcRep) / 100;

    document.getElementById('valComisVend').innerText = comissaoVend.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    
    if (tipoVenda === 'representante') {
        document.getElementById('resumoRep').classList.remove('hidden');
        document.getElementById('valComisRep').innerText = comissaoRep.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    } else {
        document.getElementById('resumoRep').classList.add('hidden');
    }
}

function limparTelaComissoes() { location.reload(); }

async function salvarComissao() {
    const btn = document.getElementById('btnRegistrar');
    const vendedora = document.getElementById('vendedora').value;
    const representante = document.getElementById('nomeRepresentante').value;
    const tipoVenda = document.getElementById('tipoVenda').value;
    const statusDiv = 'status';

    if (!vendedora) { Swal.fire('Atenção', 'Selecione a vendedora!', 'warning'); return; }
    if (tipoVenda === 'representante' && !representante) { Swal.fire('Atenção', 'Selecione o Representante!', 'warning'); return; }

    btn.disabled = true;
    exibirStatus(statusDiv, "🚀 Enviando para a Planilha Geral...", "#e2e3e5", "#333");

    Swal.fire({
        title: 'Salvando...',
        html: 'Por favor, aguarde enquanto a venda é registrada.',
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
    });

    // Otimização de DOM: Buscando valores apenas uma vez
    const valorNota = parseFloat(document.getElementById('valorNota').value) || 0;
    const valorFora = parseFloat(document.getElementById('valorFora').value) || 0;
    const porcVend = parseFloat(document.getElementById('porcVend').value) || 0;
    const totalNfVal = valorNota + valorFora;
    const nomeCliente = document.getElementById('cliente').value || "CLIENTE NÃO INFORMADO";
    const divisao = document.getElementById('divisao').value;
    const empresaSelecionada = document.getElementById('empresaSelecionada').value;

    const payload = {
        empresa: empresaSelecionada,
        nf: "AGUARDANDO", 
        pedido: "WEB-" + Math.floor(Math.random() * 9000),
        dataEmissao: new Date().toLocaleDateString('pt-BR'),
        cfop: "5102",
        razaoSocial: nomeCliente,
        statusInterno: "Impressa",
        totalNf: valorNota,
        fcp: 0, icmsSt: 0, totalOs: 0,
        totalPedido: totalNfVal,
        desconto: 0,
        mes: new Date().toLocaleString('pt-BR', { month: 'long' }).toUpperCase(),
        semana: "1",
        ano: new Date().getFullYear(),
        vendedora: vendedora,
        representante: tipoVenda === 'representante' ? representante : "VENDA DIRETA",
        divisao: divisao,
        porcentagemComissao: porcVend,
        valorComissao: totalNfVal * (porcVend / 100),
        tipoCobranca: "A DEFINIR",
        desconto2: 0,
        vencimento: "",
        statusFinanceiro: "Pendente" 
    };

    try {
        await fetch(URL_COMISSOES, {
            method: "POST",
            mode: "no-cors", // Ignora a política de CORS para o envio
            cache: "no-cache",
            headers: { "Content-Type": "text/plain" }, // text/plain garante que o envio não será bloqueado
            body: JSON.stringify(payload)
        });

        // Como usamos 'no-cors', a resposta é "opaca" e não dá pra ler o retorno JSON do Google.
        // Portanto, assumimos sucesso no envio!
        exibirStatus(statusDiv, "✅ Venda registrada com sucesso!", "#d4edda", "#155724");
        Swal.fire('Sucesso!', 'Venda registrada com sucesso!', 'success');
        
        // Atualiza Dashboard na sessão
        const valorComissao = totalNfVal * (porcVend / 100);
        totalVendidoSessao += totalNfVal;
        totalComissaoSessao += valorComissao;
        
        // Atualiza os novos Cards de KPI
        if (document.getElementById('kpiFaturamento')) {
            document.getElementById('kpiFaturamento').innerText = totalVendidoSessao.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        }
        let kpiPed = document.getElementById('kpiPedidos');
        if (kpiPed) {
            kpiPed.innerText = (parseInt(kpiPed.innerText) || 0) + 1;
        }

        // Adiciona ao Histórico
        const hora = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const nomeCliAbrev = nomeCliente.length > 15 ? nomeCliente.substring(0, 15) + "..." : nomeCliente;
        const item = `<div class="history-item"><span>${nomeCliAbrev} - ${totalNfVal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</span> <span>${hora} ✅</span></div>`;
        document.getElementById('listaHistoricoComissoes').innerHTML += item;

        // Limpa campos de valores para a próxima venda!
        document.getElementById('valorNota').value = "";
        document.getElementById('valorFora').value = "0.00";
        document.getElementById('cliente').value = "";
        calcularComissaoRealTime(); // Zera o preview na tela

    } catch (e) {
        exibirStatus(statusDiv, "❌ Erro na comunicação: " + e.message, "#f8d7da", "#721c24");
        Swal.fire('Erro!', 'Falha ao registrar a venda.', 'error');
        console.error(e);
    } finally {
        btn.disabled = false;
    }
}

async function consultarFornecedor() {
    const input = document.getElementById('cnpjInput');
    const cnpjLimpo = input.value.replace(/\D/g, '');
    const cnpjFormatado = input.value; 
    if (cnpjLimpo.length !== 14) { Swal.fire('Atenção', 'CNPJ incompleto!', 'warning'); return; }

    exibirStatus('statusFornecedor', "🔍 Consultando bases...", "#e2e3e5", "#383d41");
    document.getElementById('resultadoFornecedor').style.display = 'none';
    dadosExtras = {}; 

    exibirStatus('statusFornecedor', "🔍 Verificando se já existe...", "#fff3cd", "#856404");
    const jaExiste = await verificarDuplicidade(cnpjFormatado, cnpjLimpo);
    if (jaExiste) {
        exibirStatus('statusFornecedor', "⚠️ CNPJ JÁ CADASTRADO!", "#f8d7da", "#721c24");
        Swal.fire('Aviso', '⚠️ ATENÇÃO: Este CNPJ já consta no banco de dados!', 'info');
        return; 
    }

    try {
        const bResp = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cnpjLimpo}`);
        if (bResp.ok) { 
            const bData = await bResp.json(); 
            dadosExtras = {
                ibge: bData.codigo_municipio_ibge || "",
                cep: bData.cep || "",
                logradouro: bData.logradouro || "",
                numero: bData.numero || "",
                bairro: bData.bairro || ""
            };
        }
    } catch(e) { console.log("BrasilAPI offline"); }

    const script = document.createElement('script');
    script.src = `https://receitaws.com.br/v1/cnpj/${cnpjLimpo}?callback=callbackReceita`;
    document.body.appendChild(script);

    window.callbackReceita = (r) => {
        if (r.status === "ERROR") { Swal.fire('Erro', r.message, 'error'); return; }
        
        dadosEmpresa = {
            status: r.situacao || "ATIVO", 
            cnpj: cnpjLimpo, 
            razao_social: r.nome,
            cidade: `${r.municipio} - ${r.uf}`, 
            cod_municipio: dadosExtras.ibge || r.ibge || "N/A", 
            telefone: r.telefone || "",
            cep: r.cep || dadosExtras.cep || "N/A",
            endereco: `${r.logradouro || dadosExtras.logradouro || ''}, ${r.numero || dadosExtras.numero || ''}`,
            bairro: r.bairro || dadosExtras.bairro || "N/A"
        };
        
        document.getElementById('resRazao').innerText = r.nome;
        document.getElementById('resCnpj').innerText = cnpjLimpo;
        document.getElementById('resCidade').innerText = dadosEmpresa.cidade;
        document.getElementById('resTel').value = r.telefone; 
        
        document.getElementById('resCodMun').value = dadosEmpresa.cod_municipio;
        document.getElementById('resEmail').value = r.email ? r.email.toLowerCase() : "";
        
        document.getElementById('resultadoFornecedor').style.display = 'block';
        exibirStatus('statusFornecedor', "✅ Dados carregados com sucesso!", "#d4edda", "#155724");
        document.body.removeChild(script);
    };
}

function limparTelaFornecedores() {
    document.getElementById('cnpjInput').value = "";
    document.getElementById('resultadoFornecedor').style.display = 'none';
    document.getElementById('statusFornecedor').style.display = 'none';
    document.getElementById('resIE').value = "";
    document.getElementById('resTel').value = "";
    document.getElementById('resObs').value = "";
    document.getElementById('resEmail').value = "";
    document.getElementById('resCodMun').value = "";
    document.getElementById('resDivisao').value = "100% - SEM DIVISÃO";
}

async function salvarFornecedorNaPlanilha() {
    const btn = document.getElementById('btnSalvarFornecedor');
    const statusDiv = 'statusFornecedor'; 
    
    if (!document.getElementById('resIE').value.trim()) { 
        exibirStatus(statusDiv, "❌ IE Obrigatória!", "#f8d7da", "#721c24"); 
        return; 
    }

    btn.disabled = true;
    exibirStatus(statusDiv, "🚀 Verificando duplicidade e salvando...", "#e2e3e5", "#333");

    Swal.fire({
        title: 'Salvando Fornecedor...',
        html: 'Enviando dados para a planilha.',
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
    });

    const payload = {
        cnpj: document.getElementById('resCnpj').innerText,
        razao_social: document.getElementById('resRazao').innerText,
        cidade: document.getElementById('resCidade').innerText,
        telefone: document.getElementById('resTel').value,
        cep: dadosEmpresa.cep,
        endereco: dadosEmpresa.endereco,
        bairro: dadosEmpresa.bairro,
        ie_status: "MANUAL", 
        email: document.getElementById('resEmail').value || "N/A",
        cod_municipio: document.getElementById('resCodMun').value || "N/A",
        ie: document.getElementById('resIE').value.trim(),
        divisao: document.getElementById('resDivisao').value,
        obs: document.getElementById('resObs').value || "N/A",
        status: dadosEmpresa.status
    };

    try {
        const response = await fetch(URL_FORNECEDORES, { 
            method: 'POST',
            body: JSON.stringify(payload) 
        });

        const resultado = await response.text(); 

        if (resultado.includes("Duplicado")) {
            exibirStatus(statusDiv, "⚠️ Este CNPJ já está cadastrado na planilha!", "#fff3cd", "#856404");
            Swal.fire('Atenção', 'Este CNPJ já está cadastrado!', 'warning');
        } else {
            exibirStatus(statusDiv, "✅ Gravado com sucesso!", "#d4edda", "#155724");
            Swal.fire('Sucesso!', 'Fornecedor cadastrado com sucesso!', 'success');
            
            const hora = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            const item = `<div class="history-item"><span>${payload.razao_social.substring(0, 20)}...</span> <span>${hora} ✅</span></div>`;
            document.getElementById('listaHistorico').innerHTML += item;

            setTimeout(limparTelaFornecedores, 2000); 
        }
    } catch (e) { 
        exibirStatus(statusDiv, "❌ Erro ao conectar com o servidor.", "#f8d7da", "#721c24"); 
        Swal.fire('Erro!', 'Falha ao salvar o fornecedor.', 'error');
    } finally { 
        btn.disabled = false; 
    }
}

async function verificarDuplicidade(cnpjFmt, cnpjLimpo) {
    return new Promise((resolve) => {
        const sheetId = "1_UIvezU3eh5HQ98ttIXsViCCsY2opGwNOfZbv4SVFfc"; 
        const query = "SELECT *";
        // Truque JSONP para ignorar bloqueio de arquivo local
        const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json;responseHandler:callbackDuplicidade&gid=0&tq=${encodeURIComponent(query)}&nocache=${new Date().getTime()}`;

        window.callbackDuplicidade = (data) => {
            document.getElementById('scriptGvizDup')?.remove();
            delete window.callbackDuplicidade;
            
            if (!data.table || !data.table.rows) { resolve(false); return; }

            for (let row of data.table.rows) {
                if (row.c && row.c[1]) {
                    let valorNaPlanilha = row.c[1].v ? String(row.c[1].v) : "";
                    let valorLimpoPlanilha = valorNaPlanilha.replace(/\D/g, '');
                    
                    if (valorLimpoPlanilha === cnpjLimpo && valorLimpoPlanilha.length > 10) { resolve(true); return; }
                }
            }
            resolve(false);
        };

        const script = document.createElement('script');
        script.id = 'scriptGvizDup';
        script.src = url;
        script.onerror = () => {
            document.getElementById('scriptGvizDup')?.remove();
            delete window.callbackDuplicidade;
            console.warn("Não foi possível verificar duplicidade.");
            resolve(false);
        };
        document.body.appendChild(script);
    });
}

function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('darkMode', isDark);
    
    const btn = document.getElementById('btnDarkMode');
    if (isDark) {
        btn.classList.replace('fa-moon', 'fa-sun');
        btn.style.color = '#ffd700'; // Sol amarelado
    } else {
        btn.classList.replace('fa-sun', 'fa-moon');
        btn.style.color = 'white';
    }

    // Atualiza cores dos gráficos se estiverem na tela
    Chart.defaults.color = isDark ? '#e0e0e0' : '#666';
    if (chartMensal) {
        chartMensal.options.scales.x.grid.color = isDark ? '#444' : '#eee';
        chartMensal.options.scales.y.grid.color = isDark ? '#444' : '#eee';
        chartMensal.update();
    }
    if (chartVendedoras) chartVendedoras.update();
    
    if (!document.getElementById('telaDashboard').classList.contains('hidden')) carregarDashboardReal();
}

// --- INICIALIZAÇÕES ---
document.addEventListener("DOMContentLoaded", () => {
    mudarCorEquipe();

    // Máscaras (IMask)
    const cnpjInput = document.getElementById('cnpjInput');
    if (cnpjInput) IMask(cnpjInput, { mask: '00.000.000/0000-00' });

    const resTel = document.getElementById('resTel');
    if (resTel) IMask(resTel, { mask: '(00) 00000-0000' });

    // Aplica o Dark Mode caso o usuário já tenha ativado antes
    if(localStorage.getItem('darkMode') === 'true') {
        document.body.classList.add('dark-mode');
        const btn = document.getElementById('btnDarkMode');
        if (btn) {
            btn.classList.replace('fa-moon', 'fa-sun');
            btn.style.color = '#ffd700';
        }
    }

    // Define o mês atual no filtro do Dashboard ao abrir a página
    const mesAtualStr = new Date().toLocaleString('pt-BR', { month: 'long' }).toUpperCase();
    const filtroMes = document.getElementById('filtroMesDash');
    if (filtroMes) filtroMes.value = mesAtualStr;

    const anoAtualStr = new Date().getFullYear().toString();
    const filtroAno = document.getElementById('filtroAnoDash');
    if (filtroAno) filtroAno.value = anoAtualStr;

    // Registra o Service Worker (Aplicativo de Celular / PWA)
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js').then(() => {
            console.log('App instalado com sucesso!');
        }).catch(err => console.log('Erro no PWA:', err));
    }
});

async function carregarDashboardReal() {
    // Verifica se você configurou o ID da planilha lá em cima
    if(ID_PLANILHA_COMISSOES === "COLOQUE_AQUI_O_ID_DA_PLANILHA_DE_COMISSOES") {
        console.warn("⚠️ O Dashboard Real precisa do ID da Planilha de Comissões para funcionar.");
        return;
    }

    const vendedoraLogada = document.getElementById('vendedora').value;
    
    const userLogadoUpper = (usuarioLogado || "").toUpperCase();
    const superAdmins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE'];
    const isSuperAdmin = superAdmins.includes(userLogadoUpper);
    const isAdmin = isSuperAdmin || ['RENATA', 'CAROL'].includes(userLogadoUpper);
    if (!vendedoraLogada && !isAdmin) return;

    const filtroMes = document.getElementById('filtroMesDash');
    const mesAtual = filtroMes ? filtroMes.value : new Date().toLocaleString('pt-BR', { month: 'long' }).toUpperCase();

    const filtroAno = document.getElementById('filtroAnoDash');
    const anoAtual = filtroAno ? filtroAno.value : new Date().getFullYear().toString();

    // Pega as Datas dos Novos Filtros Avançados
    const fInicio = document.getElementById('filtroDataInicio') ? document.getElementById('filtroDataInicio').value : "";
    const fFim = document.getElementById('filtroDataFim') ? document.getElementById('filtroDataFim').value : "";
    const fVend = document.getElementById('filtroVendedoraDash') ? document.getElementById('filtroVendedoraDash').value.toUpperCase() : "TODAS";
    let valInicio = fInicio ? new Date(fInicio + "T00:00:00") : null;
    let valFim = fFim ? new Date(fFim + "T23:59:59") : null;
    
    // Usando JSONP (responseHandler) para ignorar o bloqueio de CORS ao usar o sistema no PC (file:///)
    const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_COMISSOES}/gviz/tq?tqx=out:json;responseHandler:callbackDashReal&nocache=${new Date().getTime()}`;

    // Efeito visual enquanto carrega
    if (document.getElementById('kpiFaturamento')) {
        document.getElementById('kpiFaturamento').innerText = "⏳...";
        document.getElementById('kpiPedidos').innerText = "⏳...";
        document.getElementById('kpiPendente').innerText = "⏳...";
        document.getElementById('kpiLiquido').innerText = "⏳...";
    }

    window.callbackDashReal = (data) => {
        document.getElementById('scriptGvizDash')?.remove();
        delete window.callbackDashReal;
        
        // Pela ordem que você salva os dados no Apps Script, as posições das colunas começam em 0:
        let colMes = 13;      // Mês
        let colVend = 16;     // Vendedora
        let colCli = 5;       // Razão Social do Cliente
        let colTotal = 11;    // Total do Pedido
        let colComissao = 25; // Valor da Comissão (a última da sua lista)

        if (!data.table || !data.table.rows) {
            renderizarDashAvancado([], mesAtual, anoAtual, colVend, colTotal, 0, colCli, vendedoraLogada, valInicio, valFim, fVend, colComissao);
            return;
        }

        // Popula Select de Vendedoras apenas se o usuário for Administrador
        const selectVend = document.getElementById('filtroVendedoraDash');
        
        // Defina aqui as integrantes de cada equipe para a trava funcionar:
        const equipeRenata = ['RENATA', 'VENDEDORA 1', 'VENDEDORA 2']; 
        const equipeCarol  = ['CAROL', 'VENDEDORA 3', 'VENDEDORA 4', 'VENDEDORA 5'];

        if (isAdmin && selectVend) {
            selectVend.style.display = 'inline-block';
            let vendedorasUnicas = new Set();
            data.table.rows.forEach(r => {
                if(r.c && r.c[colVend] && r.c[colVend].v) {
                    let v = String(r.c[colVend].v).toUpperCase().trim();
                    if (isSuperAdmin) vendedorasUnicas.add(v);
                    else if (userLogadoUpper === 'RENATA' && equipeRenata.includes(v)) vendedorasUnicas.add(v);
                    else if (userLogadoUpper === 'CAROL' && equipeCarol.includes(v)) vendedorasUnicas.add(v);
                }
            });
            let currentVal = selectVend.value;
            if (selectVend.options.length <= 1 || selectVend.options.length !== vendedorasUnicas.size + 1) {
                selectVend.innerHTML = '<option value="TODAS">Vendedoras (Todas)</option>';
                Array.from(vendedorasUnicas).sort().forEach(v => {
                    let opt = document.createElement('option'); opt.value = v; opt.innerText = v; selectVend.appendChild(opt);
                });
                if (vendedorasUnicas.has(currentVal) || currentVal === "TODAS") selectVend.value = currentVal;
            }
        } else if (selectVend) {
            selectVend.style.display = 'none';
        }

        let somaTotalVendas = 0;
        let somaComissao = 0;
        let countPedidos = 0;
        let comissaoPendente = 0;
        let comissaoLiquida = 0;
        let clientesMemoria = new Set(); // Memória para o Autocompletar

        for (let row of data.table.rows) {
            if(!row.c) continue;
            
            let rowMes = row.c[colMes] && row.c[colMes].v ? String(row.c[colMes].v).toUpperCase().trim() : "";
            let rowVend = row.c[colVend] && row.c[colVend].v ? String(row.c[colVend].v).toUpperCase().trim() : "";
            
            let dataEmissaoStr = row.c[3] && row.c[3].f ? String(row.c[3].f) : (row.c[3] && row.c[3].v ? String(row.c[3].v) : "");
            let rowAno = dataEmissaoStr.match(/\d{4}/) ? dataEmissaoStr.match(/\d{4}/)[0] : new Date().getFullYear().toString();

            // Converter string em Date do JS para avaliar o Período
            let rowDate = null;
            if (dataEmissaoStr.match(/\d{2}\/\d{2}\/\d{4}/)) {
                let parts = dataEmissaoStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
                rowDate = new Date(`${parts[3]}-${parts[2]}-${parts[1]}T12:00:00`);
            } else if (dataEmissaoStr.startsWith('Date')) {
                let m = dataEmissaoStr.match(/Date\((\d+),\s*(\d+),\s*(\d+)/);
                if (m) rowDate = new Date(m[1], m[2], m[3]);
            }

            // Aplica Lógica de Filtro Avançado
            let isDateValid = (valInicio && valFim && rowDate) ? (rowDate >= valInicio && rowDate <= valFim) : (rowMes === mesAtual && rowAno === anoAtual);
            
            let isVendaValida = false;
            if (isSuperAdmin) {
                isVendaValida = true; // Diretoria vê tudo
            } else if (userLogadoUpper === 'RENATA') {
                isVendaValida = equipeRenata.includes(rowVend);
            } else if (userLogadoUpper === 'CAROL') {
                isVendaValida = equipeCarol.includes(rowVend);
            } else {
                isVendaValida = (rowVend === userLogadoUpper);
            }

            if (fVend !== "TODAS" && rowVend !== fVend) isVendaValida = false;

            // Coleta o nome de todos os clientes da empresa e adiciona na memória sem repetir
            let rowCli = row.c[colCli] && row.c[colCli].v ? String(row.c[colCli].v).trim() : "";
            if (rowCli && rowCli !== "N/A") clientesMemoria.add(rowCli);

            if (isDateValid && isVendaValida) {
                let rawTotal = row.c[colTotal] ? row.c[colTotal].v : 0;
                let valTotal = typeof rawTotal === 'number' ? rawTotal : parseFloat(String(rawTotal).replace(',', '.')) || 0;
                let rawCom = row.c[colComissao] ? row.c[colComissao].v : 0;
                let valComissao = typeof rawCom === 'number' ? rawCom : parseFloat(String(rawCom).replace(',', '.')) || 0;
                
                somaTotalVendas += valTotal;
                somaComissao += valComissao;
                countPedidos++;

                // Varre a linha pra saber se a comissão já foi paga
                let isLiquido = row.c.some(c => c && typeof c.v === 'string' && c.v.toUpperCase() === 'LIQUIDO');
                if (isLiquido) comissaoLiquida += valComissao;
                else comissaoPendente += valComissao; // Padrão é pendente
            }
        }

        // Atualiza os valores na tela conforme o mês selecionado no filtro
        totalVendidoSessao = somaTotalVendas;
        totalComissaoSessao = somaComissao;
        
        document.getElementById('kpiFaturamento').innerText = totalVendidoSessao.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        document.getElementById('kpiPedidos').innerText = countPedidos;
        document.getElementById('kpiPendente').innerText = comissaoPendente.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        document.getElementById('kpiLiquido').innerText = comissaoLiquida.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

        // Popula a "listinha" do Autocomplete no campo Cliente da tela de Comissões
        const datalistCli = document.getElementById('listaClientesDatalist');
        if (datalistCli) {
            datalistCli.innerHTML = "";
            Array.from(clientesMemoria).sort().forEach(cli => {
                let opt = document.createElement('option'); opt.value = cli; datalistCli.appendChild(opt);
            });
        }

        // Envia os dados para renderizar os gráficos E a tabela
        renderizarDashAvancado(data.table.rows, mesAtual, anoAtual, colVend, colTotal, 0, colCli, vendedoraLogada, valInicio, valFim, fVend, colComissao); 
    };

    const script = document.createElement('script');
    script.id = 'scriptGvizDash';
    script.src = url;
    script.onerror = () => {
        document.getElementById('scriptGvizDash')?.remove();
        delete window.callbackDashReal;
        console.warn("Erro ao puxar dados do Dashboard.");
        if (!chartMensal) renderizarDashAvancado([], mesAtual, anoAtual, 16, 11, 0, 5, vendedoraLogada, valInicio, valFim, fVend, 25);
    };
    document.body.appendChild(script);
}

function renderizarDashAvancado(rows, mesAtual, anoAtual, colVend, colTotal, colEmpresa, colCli, vendedoraLogada, valInicio, valFim, fVend, colComissao = 25) {
    const userLogadoUpper = (usuarioLogado || "").toUpperCase();
    const superAdmins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE'];
    const isSuperAdmin = superAdmins.includes(userLogadoUpper);
    const isAdmin = isSuperAdmin || ['RENATA', 'CAROL'].includes(userLogadoUpper);
    
    const equipeRenata = ['RENATA', 'VENDEDORA 1', 'VENDEDORA 2']; 
    const equipeCarol  = ['CAROL', 'VENDEDORA 3', 'VENDEDORA 4', 'VENDEDORA 5'];
    
    const isDark = document.body.classList.contains('dark-mode');

    Chart.defaults.color = isDark ? '#e0e0e0' : '#666';

    const vendasPorVendedora = {};
    const vendasMensais = { 'JANEIRO':0, 'FEVEREIRO':0, 'MARÇO':0, 'ABRIL':0, 'MAIO':0, 'JUNHO':0, 'JULHO':0, 'AGOSTO':0, 'SETEMBRO':0, 'OUTUBRO':0, 'NOVEMBRO':0, 'DEZEMBRO':0 };
    const vendasPorCliente = {};
    let ultimasVendas = [];
    let reciboLinhas = "";
    let totalComissaoRecibo = 0;

    for (let row of rows) {
        if(!row.c) continue;
        let rowMes = row.c[13] && row.c[13].v ? String(row.c[13].v).toUpperCase().trim() : "";

        let dataEmissao = row.c[3] && row.c[3].f ? String(row.c[3].f) : (row.c[3] && row.c[3].v ? String(row.c[3].v).substring(0, 10) : "N/A");
        let vendedora = row.c[colVend] && row.c[colVend].v ? String(row.c[colVend].v).toUpperCase() : "N/A";
        let empresa = row.c[colEmpresa] && row.c[colEmpresa].v ? String(row.c[colEmpresa].v).toUpperCase() : "MISS RÔSE";
        let cliente = row.c[colCli] && row.c[colCli].v ? String(row.c[colCli].v) : "N/A";
        let rawTotal = row.c[colTotal] ? row.c[colTotal].v : 0;
        let total = typeof rawTotal === 'number' ? rawTotal : parseFloat(String(rawTotal).replace(',', '.')) || 0;
        let rowAno = dataEmissao.match(/\d{4}/) ? dataEmissao.match(/\d{4}/)[0] : new Date().getFullYear().toString();

        if (total > 0) {
            // Soma para o gráfico da Linha APENAS no ano selecionado (Mesmo com datas livres, não quebramos o resumo anual!)
            if (rowAno === anoAtual && vendasMensais[rowMes] !== undefined) vendasMensais[rowMes] += total;
            
            let rowDate = null;
            if (dataEmissao.match(/\d{2}\/\d{2}\/\d{4}/)) {
                let parts = dataEmissao.match(/(\d{2})\/(\d{2})\/(\d{4})/);
                rowDate = new Date(`${parts[3]}-${parts[2]}-${parts[1]}T12:00:00`);
            } else if (dataEmissao.startsWith('Date')) {
                let m = dataEmissao.match(/Date\((\d+),\s*(\d+),\s*(\d+)/);
                if (m) rowDate = new Date(m[1], m[2], m[3]);
            }

            let isDateValid = (valInicio && valFim && rowDate) ? (rowDate >= valInicio && rowDate <= valFim) : (rowMes === mesAtual && rowAno === anoAtual);
            
            let isVendaValida = false;
            if (isSuperAdmin) {
                isVendaValida = true;
            } else if (userLogadoUpper === 'RENATA') {
                isVendaValida = equipeRenata.includes(vendedora);
            } else if (userLogadoUpper === 'CAROL') {
                isVendaValida = equipeCarol.includes(vendedora);
            } else {
                isVendaValida = (vendedora === userLogadoUpper);
            }

            if (fVend !== "TODAS" && vendedora !== fVend) isVendaValida = false;

            if (isDateValid && isVendaValida) {
                let isLiquido = row.c.some(c => c && typeof c.v === 'string' && c.v.toUpperCase() === 'LIQUIDO');
                let statusBadge = isLiquido ? 
                    '<span style="background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: bold;"><i class="fas fa-check"></i> Líquido</span>' : 
                    '<span style="background: #fff3cd; color: #856404; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: bold;"><i class="fas fa-clock"></i> Pendente</span>';

                let rawCom = row.c[colComissao] ? row.c[colComissao].v : 0;
                let valComissao = typeof rawCom === 'number' ? rawCom : parseFloat(String(rawCom).replace(',', '.')) || 0;
                totalComissaoRecibo += valComissao;

                reciboLinhas += `
                    <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${dataEmissao}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${cliente.length > 25 ? cliente.substring(0, 25) + '...' : cliente}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: right;">R$ ${total.toFixed(2)}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #ddd; color: #2e7d32; font-weight: bold; text-align: right;">R$ ${valComissao.toFixed(2)}</td>
                    </tr>
                `;

                vendasPorVendedora[vendedora] = (vendasPorVendedora[vendedora] || 0) + total;
                
                vendasPorCliente[cliente] = (vendasPorCliente[cliente] || 0) + total;
                    ultimasVendas.unshift(`<tr>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee;">${dataEmissao}</td>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee; font-weight: bold;">${vendedora}</td>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee;">${empresa}</td>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee;">${cliente.length > 25 ? cliente.substring(0, 25) + '...' : cliente}</td>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee; color: #2e7d32; font-weight: bold;">R$ ${total.toFixed(2)}</td>
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee;">${statusBadge}</td>
                    </tr>`);
            }
        }
    }

    // Garante que o gráfico desenhe os eixos mesmo se não houver vendas
    if(Object.keys(vendasPorVendedora).length === 0) vendasPorVendedora["Sem Vendas"] = 0;

    // Preenche a tabela apenas com as últimas 50 vendas do mês para a tela não travar
    const tbody = document.getElementById('tabelaVendasDash');
    if(tbody) tbody.innerHTML = ultimasVendas.slice(0, 50).join('') || "<tr><td colspan='6' style='text-align: center; padding: 20px;'>Nenhuma venda encontrada.</td></tr>";

    const tbodyRecibo = document.getElementById('reciboTabelaBody');
    if (tbodyRecibo) {
        tbodyRecibo.innerHTML = reciboLinhas || "<tr><td colspan='4' style='text-align: center; padding: 20px;'>Nenhuma venda encontrada no período.</td></tr>";
        document.getElementById('reciboTotal').innerText = totalComissaoRecibo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        
        let nomeRecibo = (fVend && fVend !== "TODAS") ? fVend : (isAdmin ? "MÚLTIPLAS VENDEDORAS" : userLogadoUpper);
        document.getElementById('reciboVendedora').innerText = nomeRecibo;
        
        let dtInicioForm = document.getElementById('filtroDataInicio') ? document.getElementById('filtroDataInicio').value : "";
        let dtFimForm = document.getElementById('filtroDataFim') ? document.getElementById('filtroDataFim').value : "";
        let periodoRecibo = (dtInicioForm && dtFimForm) ? `${dtInicioForm.split('-').reverse().join('/')} até ${dtFimForm.split('-').reverse().join('/')}` : `${mesAtual} / ${anoAtual}`;
        document.getElementById('reciboMes').innerText = periodoRecibo;
        
        // Preenche a data de geração no formato DD/MM/AAAA HH:MM
        document.getElementById('pdfDataGeracao').innerText = new Date().toLocaleString('pt-BR');
    }

    // RENDERIZAR RANKING DE CLIENTES (TOP 5)
    const topClientes = Object.entries(vendasPorCliente).sort((a, b) => b[1] - a[1]).slice(0, 5);
    let topClientesHTML = "";
    topClientes.forEach((cli, index) => {
        let medalha = index === 0 ? "🥇" : index === 1 ? "🥈" : index === 2 ? "🥉" : "🔹";
        topClientesHTML += `
            <div style="display: flex; justify-content: space-between; padding: 12px 15px; border-bottom: 1px solid ${isDark ? '#444' : '#eee'};">
                <span style="font-weight: bold; color: var(--primary); font-size: 13px;">${medalha} ${cli[0].length > 18 ? cli[0].substring(0,18)+'...' : cli[0]}</span>
                <span style="font-weight: bold; font-size: 13px;">R$ ${cli[1].toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</span>
            </div>
        `;
    });
    
    if (topClientes.length === 0) topClientesHTML = `<div style="padding: 20px; text-align: center; color: ${isDark ? '#aaa' : '#666'};">Nenhuma venda encontrada.</div>`;
    
    const divTopCli = document.getElementById('listaTopClientesDash');
    if (divTopCli) divTopCli.innerHTML = topClientesHTML;

    // Prepara dados do Gráfico Mensal (ignorando meses no futuro que estão com 0)
    const labelsMensal = []; const dataMensal = [];
    for (let m in vendasMensais) {
        if (vendasMensais[m] > 0 || labelsMensal.length > 0) {
            labelsMensal.push(m.substring(0,3)); // Pega as 3 primeiras letras (JAN, FEV)
            dataMensal.push(vendasMensais[m]);
        }
    }
    if(labelsMensal.length === 0) { labelsMensal.push(mesAtual.substring(0,3)); dataMensal.push(0); } // Garantia pra não ficar em branco

    const ctxMensal = document.getElementById('graficoMensal').getContext('2d');
    if (chartMensal) chartMensal.destroy();
    chartMensal = new Chart(ctxMensal, {
        type: 'line',
        data: { labels: labelsMensal, datasets: [{ label: 'Faturamento Anual (R$)', data: dataMensal, borderColor: '#d81b60', backgroundColor: 'rgba(216, 27, 96, 0.1)', fill: true, tension: 0.4 }] },
        options: { 
            responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } },
            scales: {
                x: { grid: { color: isDark ? '#444' : '#eee' } },
                y: { grid: { color: isDark ? '#444' : '#eee' } }
            }
        }
    });

    const CORES_ROSA = ['#d81b60', '#ec407a', '#f06292', '#f48fb1', '#f8bbd0', '#8e44ad', '#2980b9'];
    const ctxVend = document.getElementById('graficoVendedoras').getContext('2d');
    if (chartVendedoras) chartVendedoras.destroy();
    chartVendedoras = new Chart(ctxVend, {
        type: 'doughnut',
        data: { labels: Object.keys(vendasPorVendedora), datasets: [{ data: Object.values(vendasPorVendedora), backgroundColor: CORES_ROSA, hoverOffset: 10 }] },
        options: { 
            responsive: true, 
            maintainAspectRatio: false, 
            plugins: { 
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let total = context.dataset.data.reduce((acc, curr) => acc + curr, 0);
                            let calcPorc = ((context.raw / total) * 100).toFixed(1) + '%';
                            return ' ' + context.label + ': ' + calcPorc;
                        }
                    }
                }
            } 
        }
    });
}

function exportarTabelaExcel() {
    const table = document.getElementById("tabelaExportar");
    if (!table) return;

    const mesAtual = document.getElementById('filtroMesDash') ? document.getElementById('filtroMesDash').value : "Mes";
    
    // Cria uma planilha (workbook) diretamente da tabela HTML usando SheetJS
    const wb = XLSX.utils.table_to_book(table, { sheet: "Relatório de Vendas" });
    
    // Salva o arquivo no formato Excel nativo (.xlsx)
    XLSX.writeFile(wb, `Relatorio_Vendas_${mesAtual}.xlsx`);
}

function abrirModalRecibo() {
    document.getElementById('modalRecibo').style.display = 'flex';
}

function fecharModalRecibo() {
    document.getElementById('modalRecibo').style.display = 'none';
}

async function baixarPDFReciboVisivel() {
    Swal.fire({
        title: 'Gerando PDF...',
        html: 'Aguarde um instante...',
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
    });

    const element = document.getElementById('templateReciboPDF');
    const vendedora = document.getElementById('reciboVendedora').innerText;

    const opt = {
        margin: [10, 10],
        filename: `Recibo_${vendedora}_${new Date().toLocaleDateString('pt-BR').replace(/\//g, '-')}.pdf`,
        image: { type: 'jpeg', quality: 1.0 },
        html2canvas: { 
            scale: 2, 
            useCORS: true,
            backgroundColor: '#ffffff'
        },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        await html2pdf().set(opt).from(element).save();
        Swal.fire('Sucesso!', 'Recibo baixado com sucesso.', 'success');
    } catch (error) {
        console.error("Erro ao gerar PDF:", error);
        Swal.fire('Erro!', 'Falha na geração do PDF.', 'error');
    }
}

function imprimirReciboNativo() {
    const conteudo = document.getElementById('templateReciboPDF').innerHTML;
    const telaImpressao = window.open('', '', 'width=900,height=700');
    telaImpressao.document.write(`
        <html>
        <head>
            <title>Imprimir Recibo</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 0; padding: 40px; color: black; background: white; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 25px; }
                th, td { padding: 10px; border: 1px solid #ddd; font-size: 14px; }
                th { background-color: #d81b60 !important; color: white !important; text-align: left; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                img { display: block; width: 150px; height: auto; }
            </style>
        </head>
        <body>
            ${conteudo}
        </body>
        </html>
    `);
    telaImpressao.document.close();
    telaImpressao.focus();
    setTimeout(() => {
        telaImpressao.print();
        telaImpressao.close();
    }, 500);
}