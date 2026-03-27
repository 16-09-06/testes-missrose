// URLs das suas Planilhas
const URL_COMISSOES = "https://script.google.com/macros/s/AKfycbxZO7jiev9MWxuQgAfaPdark6geTUqWAr9eYypWG_0qvqx-sVxdx6agnDu1aFIMwBL6aA/exec";
const URL_FORNECEDORES = "https://script.google.com/macros/s/AKfycbxpMXA3xWANJ8ivdoj_3ZbUV0nCXDWvJ7Ja5E6bTAdVquSImH_gDfQ9pabnwvoaZK5b/exec";
const URL_SHEET_BANCO = "https://docs.google.com/spreadsheets/d/1_UIvezU3eh5HQ98ttIXsViCCsY2opGwNOfZbv4SVFfc/edit?usp=sharing";
const URL_SHEET_LOGISTICA = "https://docs.google.com/spreadsheets/d/1inVjNncz3YdWV31iEShiYjCUkWEE0fOfkTXCwRDu98k/edit?usp=sharing";
const URL_SHEET_GERENCIAL = "https://docs.google.com/spreadsheets/d/17PYbOV8CuEwghbaDiUmJXvc1mCT7tZ55iKvkQFzTeXc/edit?usp=sharing";
const URL_LOGIN_DB = "https://script.google.com/macros/s/AKfycbyffqQQUSRWVVpyQyKyKTC5fwyEii8RzF9fFlJflwhFupAZ-QusTzhXrGSgMFEZQRHgxA/exec";

// 👇 COLOQUE A URL DO SEU NOVO SCRIPT DE PUSH AQUI 👇
const URL_PUSH_BACKEND = "https://script.google.com/macros/s/AKfycbwYIXrKAGUYam3dYYpEjHvhQA1bHDa8CYdDYE1SMcb6dewyG4XY0PR7ax_HDHFNHRHRbg/exec";
// 👇 COLOQUE SEU APP ID DO ONESIGNAL AQUI 👇
const ONESIGNAL_APP_ID = "a18338fb-0f1d-41ab-a186-42e258bb8f69";

// 👇 COLOQUE AQUI O ID DA SUA PLANILHA ONDE AS COMISSÕES SÃO SALVAS 👇
const ID_PLANILHA_COMISSOES = "17PYbOV8CuEwghbaDiUmJXvc1mCT7tZ55iKvkQFzTeXc";

// Controle de Versão do App (Mude sempre que enviar atualização)
const APP_VERSION = "1.0.9";

let usuarioLogado = localStorage.getItem('usuarioLogado');

let dadosEmpresa = {}; 
let dadosExtras = {};
let currentSheetUrl = "";
let totalVendidoSessao = 0;
let totalComissaoSessao = 0;
let chartMensal = null;
let chartVendedoras = null;
let vendasGlobaisDash = []; // Memória para o Mini-CRM

// Configuração Global do Toast (Notificações não-intrusivas)
const Toast = Swal.mixin({
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 3000,
    timerProgressBar: true,
    didOpen: (toast) => {
        toast.addEventListener('mouseenter', Swal.stopTimer)
        toast.addEventListener('mouseleave', Swal.resumeTimer)
    }
});

// Função utilitária Debounce: Otimiza performance segurando a execução de eventos repetitivos
function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// Função utilitária para converter "R$ 1.500,50" em número real (1500.50) para cálculos
function unmaskValor(valStr) {
    if(!valStr) return 0;
    return parseFloat(String(valStr).replace(/R\$\s?/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}

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

    const userUpper = user.toUpperCase();
    const superAdmins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE'];
    const isSuperAdmin = superAdmins.includes(userUpper);
    const isAdmin = isSuperAdmin || ['RENATA', 'CAROL'].includes(userUpper);

    const equipeRenata = ['RENATA', 'HOZANA', 'ISRAEL', 'ROSANGELA', 'SARA', 'VINICIUS'];
    const equipeCarol  = ['CAROL', 'ALICE', 'CHARLENE', 'HEMILLY', 'MICHELLE'];

    // Limpa as opções antigas e preenche com a equipe correta
    vendedoraSelect.innerHTML = '<option value="">Selecione...</option>';
    let vendedorasPermitidas = [];
    
    if (isSuperAdmin) {
        vendedorasPermitidas = [...new Set([...equipeRenata, ...equipeCarol, userUpper])].sort();
    } else if (userUpper === 'RENATA') {
        vendedorasPermitidas = equipeRenata;
    } else if (userUpper === 'CAROL') {
        vendedorasPermitidas = equipeCarol;
    } else {
        vendedorasPermitidas = [userUpper];
    }

    vendedorasPermitidas.forEach(v => vendedoraSelect.add(new Option(v, v)));

    vendedoraSelect.value = userUpper;
    vendedoraSelect.disabled = !isAdmin;

    // Aciona a busca do Dashboard real após identificar a vendedora
    setTimeout(carregarDashboardReal, 1000); 
    if(!window.dashInterval) {
        window.dashInterval = setInterval(carregarDashboardReal, 60000); // Atualiza os gráficos a cada 1 min
    }

    // NOVO: Configura recursos de admin
    setupAdminFeatures(user);
}

function setupAdminFeatures(user) {
    const userUpper = user.toUpperCase();
    const superAdmins = ['KAYK', 'JHONATA', 'DEBORA', 'FELIPE'];
    const isSuperAdmin = superAdmins.includes(userUpper);
    const isAdmin = isSuperAdmin || ['RENATA', 'CAROL'].includes(userUpper);

    const painelPush = document.getElementById('painelAdminPush');
    if (!painelPush) return;

    if (isAdmin) {
        painelPush.classList.remove('hidden');
        const pushTargetSelect = document.getElementById('pushTarget');
        pushTargetSelect.innerHTML = ''; // Limpa opções

        if (userUpper === 'RENATA') {
            pushTargetSelect.add(new Option('Minha Equipe (Renata)', 'equipe_renata'));
        } else if (userUpper === 'CAROL') {
            pushTargetSelect.add(new Option('Minha Equipe (Carol)', 'equipe_carol'));
        } else if (isSuperAdmin) {
            pushTargetSelect.add(new Option('Todas as Vendedoras', 'todas'));
            pushTargetSelect.add(new Option('Equipe Renata', 'equipe_renata'));
            pushTargetSelect.add(new Option('Equipe Carol', 'equipe_carol'));
        }
    } else {
        painelPush.classList.add('hidden');
    }
}

async function realizarLogin() {
    const user = document.getElementById('userLogin').value.trim().toUpperCase();
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
    const nome = document.getElementById('novoNome').value.trim().toUpperCase();
    const senhaRaw = document.getElementById('novaSenha').value;
    const btn = document.querySelector('#formCadastrar button');

    if (!nome || !senhaRaw) return Swal.fire('Atenção', 'Por favor, preencha nome e senha.', 'warning');
    
    btn.disabled = true;
    btn.innerHTML = '💾 Salvando...';
    const senhaHash = await hashPassword(senhaRaw);

    try {
        const response = await fetch(URL_LOGIN_DB, {
            method: 'POST',
            body: JSON.stringify({ acao: "cadastrar", nome: nome, senha: senhaHash })
        });
        const texto = await response.text();
        
        if (texto === "Sucesso") {
            Swal.fire({
                icon: 'success',
                title: 'Cadastro Realizado!',
                text: `O usuário ${nome} foi criado. Agora você já pode entrar no sistema.`,
                timer: 3000,
                showConfirmButton: false
            });

            document.getElementById('novoNome').value = '';
            document.getElementById('novaSenha').value = '';
            alternarAbaAuth('entrar');
            document.getElementById('userLogin').value = nome;
            document.getElementById('senhaLogin').focus();
        } else {
            Swal.fire('Erro', 'Não foi possível realizar o cadastro. O usuário pode já existir.', 'error');
        }
    } catch (e) {
        Swal.fire('Erro', 'Erro ao conectar com a planilha de acesso.', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '💾 Salvar Cadastro';
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

function mostrarSenha(inputId = 'senhaInput', iconId = 'toggleSenha') {
    const s = document.getElementById(inputId);
    const i = document.getElementById(iconId);
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
    const telaConfig = document.getElementById('telaConfig');
    const navDash = document.getElementById('navDashboard');
    const navCom = document.getElementById('navComissoes');
    const navFor = document.getElementById('navFornecedores');
    const navPlan = document.getElementById('navPlanilhas');
    const navConfig = document.getElementById('navConfig');

    if (telaDash) telaDash.classList.add('hidden');
    telaCom.classList.add('hidden');
    telaFor.classList.add('hidden');
    telaPlan.classList.add('hidden');
    if (telaConfig) telaConfig.classList.add('hidden');
    
    if (navDash) navDash.classList.remove('active');
    navCom.classList.remove('active');
    navFor.classList.remove('active');
    navPlan.classList.remove('active');
    if (navConfig) navConfig.classList.remove('active');

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
    } else if (tela === 'config') {
        if (telaConfig) telaConfig.classList.remove('hidden');
        if (navConfig) navConfig.classList.add('active');
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

const calcularComissaoDebounced = debounce(calcularComissaoRealTime, 300);

document.querySelectorAll('#valorNota, #valorFora, #porcVend, #porcRep').forEach(el => {
    el.addEventListener('input', calcularComissaoDebounced);
});

function calcularComissaoRealTime() {
    const valorNota = unmaskValor(document.getElementById('valorNota').value);
    const valorFora = unmaskValor(document.getElementById('valorFora').value);
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

function limparRascunhoVenda() {
    localStorage.removeItem('rascunhoVenda');
}

function limparTelaComissoes() {
    Swal.fire({
        title: 'Deseja limpar tudo?',
        text: "Todos os dados da venda atual serão perdidos.",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d81b60',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Sim, limpar',
        cancelButtonText: 'Cancelar'
    }).then((result) => { 
        if (result.isConfirmed) {
            limparRascunhoVenda();
            location.reload(); 
        }
    });
}

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
    const valorNota = unmaskValor(document.getElementById('valorNota').value);
    const valorFora = unmaskValor(document.getElementById('valorFora').value);
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

    // --- MODO OFFLINE (FILA DE VENDAS) ---
    if (!navigator.onLine) {
        let fila = JSON.parse(localStorage.getItem('vendasOfflineQueue') || '[]');
        fila.push(payload);
        localStorage.setItem('vendasOfflineQueue', JSON.stringify(fila));
        
        exibirStatus(statusDiv, "📴 Sem internet. Salvo na fila offline!", "#fff3cd", "#856404");
        Toast.fire({ icon: 'info', title: 'Você está offline. Venda salva na fila local!' });
        
        const hora = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const nomeCliAbrev = nomeCliente.length > 15 ? nomeCliente.substring(0, 15) + "..." : nomeCliente;
        const item = `
        <div class="history-item" style="align-items: center; color: #856404;">
            <span>${nomeCliAbrev} - ${totalNfVal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</span> 
            <div><span>${hora} 📴</span></div>
        </div>`;
        document.getElementById('listaHistoricoComissoes').innerHTML += item;

        document.getElementById('valorNota').value = "";
        document.getElementById('valorFora').value = "";
        document.getElementById('cliente').value = "";
        calcularComissaoRealTime();
        limparRascunhoVenda();
        btn.disabled = false;
        
        // Badge no botão para avisar que tem vendas pendentes
        if(!document.getElementById('badgeOffline')) {
            btn.innerHTML += ` <span id="badgeOffline" style="background: white; color: #856404; padding: 2px 6px; border-radius: 10px; font-size: 12px; margin-left: 5px;">Offline</span>`;
        }
        return;
    }

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
        Toast.fire({ icon: 'success', title: 'Venda registrada com sucesso!' });
        
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
        const payloadStr = encodeURIComponent(JSON.stringify(payload));
        const item = `
        <div class="history-item" style="align-items: center;">
            <span>${nomeCliAbrev} - ${totalNfVal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</span> 
            <div>
                <span>${hora} ✅</span>
                <button onclick="editarVendaSessao('${payloadStr}')" style="width: auto; padding: 4px 8px; margin: 0 0 0 10px; font-size: 12px; background: #ffc107; color: #333; border-radius: 4px;" title="Editar esta venda"><i class="fas fa-edit"></i></button>
            </div>
        </div>`;
        document.getElementById('listaHistoricoComissoes').innerHTML += item;

        // Limpa campos de valores para a próxima venda!
        document.getElementById('valorNota').value = "";
        document.getElementById('valorFora').value = "";
        document.getElementById('cliente').value = "";
        calcularComissaoRealTime(); // Zera o preview na tela
        limparRascunhoVenda(); // Limpa o rascunho pós sucesso

    } catch (e) {
        exibirStatus(statusDiv, "❌ Erro na comunicação: " + e.message, "#f8d7da", "#721c24");
        Swal.fire('Erro!', 'Falha ao registrar a venda.', 'error');
        console.error(e);
    } finally {
        btn.disabled = false;
    }
}

function editarVendaSessao(payloadStr) {
    const data = JSON.parse(decodeURIComponent(payloadStr));
    
    Swal.fire({
        title: 'Corrigir Lançamento?',
        text: "Os dados voltarão para o formulário. Lembre-se de avisar a gerência para ignorar o registro anterior incorreto na planilha.",
        icon: 'info',
        showCancelButton: true,
        confirmButtonColor: '#ffc107',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Sim, carregar dados',
        cancelButtonText: 'Cancelar'
    }).then((result) => {
        if (result.isConfirmed) {
            document.getElementById('empresaSelecionada').value = data.empresa || "MISS RÔSE";
            document.getElementById('vendedora').value = data.vendedora || "";
            
            if (data.representante && data.representante !== "VENDA DIRETA") {
                document.getElementById('tipoVenda').value = 'representante';
                toggleRepresentante();
                document.getElementById('nomeRepresentante').value = data.representante;
            } else {
                document.getElementById('tipoVenda').value = 'direta';
                toggleRepresentante();
            }
            
            document.getElementById('cliente').value = data.razaoSocial || "";
            document.getElementById('divisao').value = data.divisao || "100% - SEM DIVISÃO";
            document.getElementById('porcVend').value = data.porcentagemComissao || "";
            
            const valorNota = parseFloat(data.totalNf) || 0;
            const valorFora = (parseFloat(data.totalPedido) || 0) - valorNota;
            
            document.getElementById('valorNota').value = valorNota.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
            document.getElementById('valorFora').value = valorFora.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
            
            calcularComissaoRealTime();
            window.scrollTo({ top: 0, behavior: 'smooth' }); // Rola a tela para o topo
        }
    });
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
    Swal.fire({
        title: 'Limpar Cadastro?',
        text: "Você perderá os dados não salvos.",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d81b60',
        cancelButtonColor: '#6c757d',
        confirmButtonText: 'Sim, limpar',
        cancelButtonText: 'Cancelar'
    }).then((result) => {
        if (result.isConfirmed) {
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
    });
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
            Toast.fire({ icon: 'success', title: 'Fornecedor cadastrado com sucesso!' });
            
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

function mostrarPopupAtualizacao(registration) {
    Swal.fire({
        title: 'Nova Versão Disponível!',
        html: 'Uma atualização foi encontrada. Recomendamos atualizar agora para ter os recursos mais recentes.',
        icon: 'info',
        showCancelButton: true,
        confirmButtonColor: '#28a745',
        cancelButtonColor: '#6c757d',
        confirmButtonText: '<i class="fas fa-sync-alt"></i> Atualizar Agora',
        cancelButtonText: 'Depois',
        allowOutsideClick: false,
        allowEscapeKey: false
    }).then((result) => {
        if (result.isConfirmed) {
            // Pede para o novo Service Worker assumir o controle
            if (registration.waiting) {
                registration.waiting.postMessage({ action: 'skipWaiting' });
            }
        }
    });
}

async function verificarAtualizacaoManual() {
    const btn = document.getElementById('btnCheckUpdate');
    if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Verificando no servidor...'; }

    if ('serviceWorker' in navigator) {
        try {
            const registration = await navigator.serviceWorker.getRegistration();
            if (registration) {
                // Força o SW a checar o arquivo sw.js no servidor
                await registration.update();
                
                // Damos um pequeno delay para a checagem acontecer e o evento onupdatefound disparar
                setTimeout(() => {
                    if (!registration.installing) {
                        // Se não encontrou instalador novo após o update(), significa que já está na última versão
                        Toast.fire({ icon: 'success', title: 'Seu aplicativo já está na versão mais recente!' });
                    }
                    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-sync-alt"></i> Verificar Atualizações'; }
                }, 1500);
            }
        } catch (e) {
            Toast.fire({ icon: 'error', title: 'Erro ao verificar atualizações.' });
            if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-sync-alt"></i> Verificar Atualizações'; }
        }
    } else {
        if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-sync-alt"></i> Verificar Atualizações'; }
    }
}

// --- AUTO-SAVE (Salva-Vidas) ---
function salvarRascunhoVenda() {
    const campos = ['empresaSelecionada', 'equipe', 'vendedora', 'tipoVendedora', 'tipoVenda', 'nomeRepresentante', 'tipoRepresentante', 'porcRep', 'cliente', 'valorNota', 'valorFora', 'divisao', 'porcVend'];
    const rascunho = {};
    campos.forEach(id => {
        const el = document.getElementById(id);
        if(el) rascunho[id] = el.value;
    });
    localStorage.setItem('rascunhoVenda', JSON.stringify(rascunho));
}

function verificarRascunhoPendente() {
    const rascunhoSalvo = localStorage.getItem('rascunhoVenda');
    if (rascunhoSalvo) {
        const r = JSON.parse(rascunhoSalvo);
        if(r.valorNota || r.valorFora || r.cliente) { // Só avisa se tiver algo importante preenchido
            Swal.fire({
                title: 'Rascunho Encontrado',
                text: "Você deixou uma venda pela metade na última vez. Deseja continuar?",
                icon: 'info',
                showCancelButton: true,
                confirmButtonColor: '#d81b60',
                cancelButtonColor: '#6c757d',
                confirmButtonText: 'Sim, restaurar',
                cancelButtonText: 'Não, limpar'
            }).then((result) => {
                if (result.isConfirmed) {
                    Object.keys(r).forEach(id => {
                        const el = document.getElementById(id);
                        if(el) el.value = r[id];
                    });
                    toggleRepresentante();
                    calcularComissaoRealTime();
                } else {
                    limparRascunhoVenda();
                }
            });
        }
    }
}

// --- SINCRONIZAÇÃO OFFLINE QUANDO A INTERNET VOLTAR ---
async function sincronizarFilaOffline() {
    let fila = JSON.parse(localStorage.getItem('vendasOfflineQueue') || '[]');
    if(fila.length === 0) return;
    
    Toast.fire({ icon: 'info', title: `Sincronizando ${fila.length} venda(s) offline...` });
    let filaRestante = [];
    for(let payload of fila) {
        try {
            await fetch(URL_COMISSOES, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain" }, body: JSON.stringify(payload) });
        } catch(e) {
            filaRestante.push(payload); // Mantém se falhar
        }
    }
    localStorage.setItem('vendasOfflineQueue', JSON.stringify(filaRestante));
    if(filaRestante.length === 0) { Toast.fire({ icon: 'success', title: 'Todas as vendas sincronizadas!' }); const btn = document.getElementById('btnRegistrar'); if(btn) btn.innerHTML = '✅ Registrar Venda'; }
}

// --- INICIALIZAÇÕES ---
document.addEventListener("DOMContentLoaded", () => {
    mudarCorEquipe();
    atualizarBadgeNotificacoes(); // Inicia o contador de notificações

    // Máscaras (IMask)
    const cnpjInput = document.getElementById('cnpjInput');
    if (cnpjInput) IMask(cnpjInput, { mask: '00.000.000/0000-00' });

    const resTel = document.getElementById('resTel');
    if (resTel) IMask(resTel, { mask: '(00) 00000-0000' });

    // Máscaras de Dinheiro para a Tela de Comissões
    const maskDinheiro = { mask: 'R$ num', blocks: { num: { mask: Number, scale: 2, thousandsSeparator: '.', padFractionalZeros: true, normalizeZeros: true, radix: ',' } } };
    const inNota = document.getElementById('valorNota');
    if(inNota) IMask(inNota, maskDinheiro);
    const inFora = document.getElementById('valorFora');
    if(inFora) IMask(inFora, maskDinheiro);

    // Aplica o Dark Mode caso o usuário já tenha ativado antes
    if(localStorage.getItem('darkMode') === 'true') {
        document.body.classList.add('dark-mode');
        const btn = document.getElementById('btnDarkMode');
        if (btn) {
            btn.classList.replace('fa-moon', 'fa-sun');
            btn.style.color = '#ffd700';
        }
    }

    // Preenche a versão atual na tela de Configurações
    const versionDisplay = document.getElementById('appVersionDisplay');
    if (versionDisplay) versionDisplay.innerText = "Versão " + APP_VERSION;

    // Define o mês atual no filtro do Dashboard ao abrir a página
    const mesAtualStr = new Date().toLocaleString('pt-BR', { month: 'long' }).toUpperCase();
    const filtroMes = document.getElementById('filtroMesDash');
    if (filtroMes) filtroMes.value = mesAtualStr;

    const anoAtualStr = new Date().getFullYear().toString();
    const filtroAno = document.getElementById('filtroAnoDash');
    if (filtroAno) filtroAno.value = anoAtualStr;

    // Adiciona Monitoramento de Rascunho na tela de Comissões
    const formComissao = document.getElementById('telaComissoes');
    if(formComissao) {
        // Salva rascunho ao mudar de campo (rápido)
        formComissao.querySelectorAll('input, select').forEach(el => el.addEventListener('change', salvarRascunhoVenda));
        // Salva rascunho enquanto digita (com delay para não sobrecarregar)
        formComissao.querySelectorAll('input').forEach(el => el.addEventListener('input', debounce(salvarRascunhoVenda, 1000)));
    }

    window.addEventListener('online', sincronizarFilaOffline);
    sincronizarFilaOffline(); // Tenta sincronizar a fila logo na abertura do app
    setTimeout(verificarRascunhoPendente, 2000); // Dá tempo de carregar os selects de vendedora

    // Registra o Service Worker (Aplicativo de Celular / PWA)
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js').then(registration => {
            console.log('Service Worker registrado com sucesso!');

            registration.onupdatefound = () => {
                const installingWorker = registration.installing;
                if (installingWorker == null) {
                    return;
                }
                installingWorker.onstatechange = () => {
                    if (installingWorker.state === 'installed') {
                        if (navigator.serviceWorker.controller) {
                            // Neste ponto, o novo SW está instalado e esperando.
                            // É o momento perfeito para notificar o usuário.
                            console.log('Nova versão do app disponível! Por favor, atualize.');
                            mostrarPopupAtualizacao(registration);
                        } else {
                            // Conteúdo está todo em cache. App pronto para uso offline.
                            console.log('Conteúdo em cache para uso offline.');
                        }
                    }
                };
            };
        }).catch(error => {
            console.error('Erro no registro do Service Worker:', error);
        });

        // Recarrega a página uma vez que o novo SW tomou o controle
        let refreshing;
        navigator.serviceWorker.addEventListener('controllerchange', () => {
            if (refreshing) return;
            window.location.reload();
            refreshing = true;
        });
    }

    initOneSignal(); // Inicia o motor de mensagens
    // Checa o status da inscrição Push ao carregar a página
    checkPushSubscription();
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

        vendasGlobaisDash = data.table.rows || []; // Salva pro Mini-CRM

        if (!data.table || !data.table.rows) {
            renderizarDashAvancado([], mesAtual, anoAtual, colVend, colTotal, 0, colCli, vendedoraLogada, valInicio, valFim, fVend, colComissao);
            return;
        }

        // Popula Select de Vendedoras apenas se o usuário for Administrador
        const selectVend = document.getElementById('filtroVendedoraDash');
        
        // Defina aqui as integrantes de cada equipe para a trava funcionar:
        const equipeRenata = ['RENATA', 'HOZANA', 'ISRAEL', 'ROSANGELA', 'SARA', 'VINICIUS']; 
        const equipeCarol  = ['CAROL', 'ALICE', 'CHARLENE', 'HEMILLY', 'MICHELLE'];

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
        
        // Variáveis para Comparação de Mês Anterior
        const isAnoInteiro = mesAtual === "ANO INTEIRO";
        const mesesRef = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
        let idxMes = mesesRef.indexOf(mesAtual);
        let mesPrevStr = isAnoInteiro ? "" : (idxMes <= 0 ? 'DEZEMBRO' : mesesRef[idxMes - 1]);
        let anoPrevStr = isAnoInteiro ? (parseInt(anoAtual) - 1).toString() : (idxMes <= 0 ? (parseInt(anoAtual) - 1).toString() : anoAtual);
        
        let somaTotalVendasPrev = 0;
        let countPedidosPrev = 0;
        let comissaoPendentePrev = 0;
        let comissaoLiquidaPrev = 0;

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
            let isDateValid = (valInicio && valFim && rowDate) ? (rowDate >= valInicio && rowDate <= valFim) : ((isAnoInteiro || rowMes === mesAtual) && rowAno === anoAtual);
            let isDateValidPrev = (valInicio && valFim && rowDate) ? false : ((isAnoInteiro || rowMes === mesPrevStr) && rowAno === anoPrevStr); // Não calcula % em datas customizadas
            
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
            
            // Acumula os dados se a venda for do mês ANTERIOR
            if (isDateValidPrev && isVendaValida) {
                let rawTotal = row.c[colTotal] ? row.c[colTotal].v : 0;
                somaTotalVendasPrev += (typeof rawTotal === 'number' ? rawTotal : parseFloat(String(rawTotal).replace(',', '.')) || 0);
                countPedidosPrev++;
                let rawCom = row.c[colComissao] ? row.c[colComissao].v : 0;
                let valComPrev = typeof rawCom === 'number' ? rawCom : parseFloat(String(rawCom).replace(',', '.')) || 0;
                let isLiquido = row.c.some(c => c && typeof c.v === 'string' && c.v.toUpperCase() === 'LIQUIDO');
                if (isLiquido) comissaoLiquidaPrev += valComPrev; else comissaoPendentePrev += valComPrev;
            }
        }

        // Atualiza os valores na tela conforme o mês selecionado no filtro
        totalVendidoSessao = somaTotalVendas;
        totalComissaoSessao = somaComissao;
        
        document.getElementById('kpiFaturamento').innerText = totalVendidoSessao.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        document.getElementById('kpiPedidos').innerText = countPedidos;
        document.getElementById('kpiPendente').innerText = comissaoPendente.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        document.getElementById('kpiLiquido').innerText = comissaoLiquida.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

        // Lógica para desenhar as setas de crescimento (Se houver crescimento de mês atual x mês anterior)
        function getBadgeCrescimento(atual, prev) {
            if (valInicio && valFim) return ""; // Esconde setinhas se o filtro for por dias específicos
            if (prev === 0) return atual > 0 ? `<span style="color:#28a745;">(🔺 100%)</span>` : ``;
            let diff = ((atual - prev) / prev) * 100;
            if (diff > 0) return `<span style="color:#28a745;">(🔺 +${diff.toFixed(1)}%)</span>`;
            if (diff < 0) return `<span style="color:#dc3545;">(🔻 ${diff.toFixed(1)}%)</span>`;
            return `<span style="color:#888;">(⏸ 0%)</span>`;
        }
        if(document.getElementById('kpiFaturamentoPrev')) document.getElementById('kpiFaturamentoPrev').innerHTML = getBadgeCrescimento(somaTotalVendas, somaTotalVendasPrev);
        if(document.getElementById('kpiPedidosPrev')) document.getElementById('kpiPedidosPrev').innerHTML = getBadgeCrescimento(countPedidos, countPedidosPrev);
        if(document.getElementById('kpiPendentePrev')) document.getElementById('kpiPendentePrev').innerHTML = getBadgeCrescimento(comissaoPendente, comissaoPendentePrev);
        if(document.getElementById('kpiLiquidoPrev')) document.getElementById('kpiLiquidoPrev').innerHTML = getBadgeCrescimento(comissaoLiquida, comissaoLiquidaPrev);

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
    
    const equipeRenata = ['RENATA', 'HOZANA', 'ISRAEL', 'ROSANGELA', 'SARA', 'VINICIUS']; 
    const equipeCarol  = ['CAROL', 'ALICE', 'CHARLENE', 'HEMILLY', 'MICHELLE'];
    
    const isDark = document.body.classList.contains('dark-mode');

    Chart.defaults.color = isDark ? '#e0e0e0' : '#666';

    const vendasPorVendedora = {};
    const vendasPorCliente = {};
    const isAnoInteiro = mesAtual === "ANO INTEIRO";
    const faturamentoAtual = isAnoInteiro ? new Array(12).fill(0) : new Array(31).fill(0);
    const faturamentoPrev = isAnoInteiro ? new Array(12).fill(0) : new Array(31).fill(0);
    
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
            let rowDate = null;
            if (dataEmissao.match(/\d{2}\/\d{2}\/\d{4}/)) {
                let parts = dataEmissao.match(/(\d{2})\/(\d{2})\/(\d{4})/);
                rowDate = new Date(`${parts[3]}-${parts[2]}-${parts[1]}T12:00:00`);
            } else if (dataEmissao.startsWith('Date')) {
                let m = dataEmissao.match(/Date\((\d+),\s*(\d+),\s*(\d+)/);
                if (m) rowDate = new Date(m[1], m[2], m[3]);
            }

            // Tenta achar o dia/mês para o Gráfico Comparativo
            let diaVenda = 1;
            let mesVendaIdx = 0;
            if (dataEmissao.match(/\d{2}\/\d{2}\/\d{4}/)) {
                let parts = dataEmissao.split('/');
                diaVenda = parseInt(parts[0], 10);
                mesVendaIdx = parseInt(parts[1], 10) - 1;
            } else if (rowDate) {
                diaVenda = rowDate.getDate();
                mesVendaIdx = rowDate.getMonth();
            }

            let isDateValid = (valInicio && valFim && rowDate) ? (rowDate >= valInicio && rowDate <= valFim) : ((isAnoInteiro || rowMes === mesAtual) && rowAno === anoAtual);
            
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
                // Soma para a linha do Gráfico
                if (isAnoInteiro) {
                    if (mesVendaIdx >= 0 && mesVendaIdx <= 11) faturamentoAtual[mesVendaIdx] += total;
                } else {
                    if(diaVenda >= 1 && diaVenda <= 31) faturamentoAtual[diaVenda - 1] += total;
                }
                
                // ... [O resto do IF continua igual preenchendo as tabelas]
            }

            // Condição separada para popular a linha cinza (Período Anterior) do gráfico
            const mesesRef = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
            let idxM = mesesRef.indexOf(mesAtual);
            let mesPrevS = isAnoInteiro ? "" : (idxM <= 0 ? 'DEZEMBRO' : mesesRef[idxM - 1]);
            let anoPrevS = isAnoInteiro ? (parseInt(anoAtual) - 1).toString() : (idxM <= 0 ? (parseInt(anoAtual) - 1).toString() : anoAtual);
            
            let isPrevValid = isAnoInteiro ? (rowAno === anoPrevS) : (rowMes === mesPrevS && rowAno === anoPrevS);
            
            if (isVendaValida && isPrevValid) {
                if (isAnoInteiro) {
                    if (mesVendaIdx >= 0 && mesVendaIdx <= 11) faturamentoPrev[mesVendaIdx] += total;
                } else {
                    if (diaVenda >= 1 && diaVenda <= 31) faturamentoPrev[diaVenda - 1] += total;
                }
            }

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
                        <td style="padding: 10px 15px; border-bottom: 1px solid #eee;"><span class="clickable-client" onclick="abrirFichaCliente(this.textContent)">${cliente.length > 25 ? cliente.substring(0, 25) + '...' : cliente}</span></td>
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
                <span class="clickable-client" style="font-weight: bold; font-size: 13px;" onclick="abrirFichaCliente(this.textContent.replace(/^[^a-zA-Z0-9]+/, '').trim())">${medalha} ${cli[0].length > 18 ? cli[0].substring(0,18)+'...' : cli[0]}</span>
                <span style="font-weight: bold; font-size: 13px;">R$ ${cli[1].toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</span>
            </div>
        `;
    });
    
    if (topClientes.length === 0) topClientesHTML = `<div style="padding: 20px; text-align: center; color: ${isDark ? '#aaa' : '#666'};">Nenhuma venda encontrada.</div>`;
    
    const divTopCli = document.getElementById('listaTopClientesDash');
    if (divTopCli) divTopCli.innerHTML = topClientesHTML;

    // Prepara Gráfico Comparativo
    const labelsGrafico = isAnoInteiro ? ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'] : Array.from({length: 31}, (_, i) => (i + 1).toString());
    const labelAtual = isAnoInteiro ? `Vendas em ${anoAtual}` : `Vendas em ${mesAtual}`;
    const labelPrev = isAnoInteiro ? `Ano Anterior` : `Mês Anterior`;
    
    const tit = document.getElementById('tituloEvolucao');
    if (tit) tit.innerText = isAnoInteiro ? 'Evolução Mensal (Ano Atual x Ano Anterior)' : 'Evolução Diária (Mês Atual x Mês Anterior)';
    
    const ctxMensal = document.getElementById('graficoMensal').getContext('2d');
    if (chartMensal) chartMensal.destroy();
    chartMensal = new Chart(ctxMensal, {
        type: 'line',
        data: { 
            labels: labelsGrafico, 
            datasets: [
                { label: labelAtual, data: faturamentoAtual, borderColor: '#d81b60', backgroundColor: 'rgba(216, 27, 96, 0.1)', fill: true, tension: 0.4 },
                { label: labelPrev, data: faturamentoPrev, borderColor: isDark ? '#888' : '#bbb', borderDash: [5, 5], fill: false, tension: 0.4, borderWidth: 2 }
            ] 
        },
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

// --- LÓGICA DO MINI-CRM (FICHA DO CLIENTE) ---
function abrirFichaCliente(nomeClienteLimpo) {
    const nomeBusca = nomeClienteLimpo.replace(/\.\.\.$/, '').trim().toUpperCase();
    document.getElementById('fichaNomeCliente').innerText = nomeBusca;
    
    let totalGasto = 0;
    let dataUltimaCompra = null;
    let historicoHtml = "";
    
    // Varre todas as linhas puxadas da planilha buscando esse cliente
    vendasGlobaisDash.forEach(row => {
        if(!row.c) return;
        let rowCli = row.c[5] && row.c[5].v ? String(row.c[5].v).toUpperCase().trim() : "";
        if (rowCli.includes(nomeBusca)) {
            let dataEmissao = row.c[3] && row.c[3].f ? String(row.c[3].f) : (row.c[3] && row.c[3].v ? String(row.c[3].v).substring(0, 10) : "N/A");
            let rawTotal = row.c[11] ? row.c[11].v : 0;
            let valTotal = typeof rawTotal === 'number' ? rawTotal : parseFloat(String(rawTotal).replace(',', '.')) || 0;
            
            totalGasto += valTotal;
            if(!dataUltimaCompra) dataUltimaCompra = dataEmissao; // Pega a primeira que achar (está decrescente)
            
            historicoHtml += `
                <div style="display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dotted #ccc;">
                    <span>📅 ${dataEmissao}</span>
                    <strong style="color: var(--primary);">R$ ${valTotal.toFixed(2)}</strong>
                </div>
            `;
        }
    });

    document.getElementById('fichaTotalGasto').innerText = totalGasto.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    document.getElementById('fichaUltimaCompra').innerText = dataUltimaCompra || "N/A";
    document.getElementById('fichaListaCompras').innerHTML = historicoHtml || "<div>Sem registros anteriores.</div>";
    
    document.getElementById('modalFichaCliente').classList.remove('hidden');
}

function fecharFichaCliente() {
    document.getElementById('modalFichaCliente').classList.add('hidden');
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

async function compartilharReciboWhatsApp() {
    const vendedora = document.getElementById('reciboVendedora').innerText;
    const mes = document.getElementById('reciboMes').innerText;
    const total = document.getElementById('reciboTotal').innerText;

    const textoRecibo = `*Resumo de Vendas - Miss Rôse*\n` +
                        `👤 Vendedora: ${vendedora}\n` +
                        `📅 Período: ${mes}\n` +
                        `💰 Total de Comissão: ${total}\n\n` +
                        `Gerado pelo App Miss Rôse`;

    if (navigator.share) {
        try {
            await navigator.share({
                title: 'Recibo Miss Rôse',
                text: textoRecibo,
            });
        } catch (err) {
            console.log('Compartilhamento cancelado ou falhou', err);
        }
    } else {
        window.open(`https://api.whatsapp.com/send?text=${encodeURIComponent(textoRecibo)}`);
    }
}

// --- LÓGICA PARA NOTIFICAÇÕES PUSH ---

function initOneSignal() {
    window.OneSignalDeferred = window.OneSignalDeferred || [];
    OneSignalDeferred.push(async function(OneSignal) {
        await OneSignal.init({
            appId: ONESIGNAL_APP_ID,
            notifyButton: { enable: false },
        });

        // Intercepta a notificação e joga para o nosso "Sininho" do aplicativo
        OneSignal.Notifications.addEventListener('foregroundWillDisplay', function(event) {
            const notif = event.notification;
            salvarNotificacaoLocal({ title: notif.title, body: notif.body });
            Toast.fire({ icon: 'info', title: 'Você tem uma nova mensagem!' });
        });
    });
}

async function subscribeUserToPush() {
    const pushStatus = document.getElementById('pushStatus');
    const pushButton = document.getElementById('btnHabilitarPush');

    window.OneSignalDeferred.push(async function(OneSignal) {
        pushStatus.innerText = 'Solicitando permissão...';
        await OneSignal.Notifications.requestPermission();
        
        if (OneSignal.Notifications.permission === true) {
            pushStatus.innerText = 'Você está inscrito para receber notificações!';
            pushButton.innerHTML = '<i class="fas fa-check-circle"></i> Inscrito';
            pushButton.disabled = true;
            Toast.fire({ icon: 'success', title: 'Notificações habilitadas!' });
            
            tagOneSignalUser(usuarioLogado); // Classifica a vendedora na equipe dela
        } else {
            pushStatus.innerText = 'Permissão negada pelo navegador.';
            pushButton.disabled = false;
        }
    });
}

function tagOneSignalUser(user) {
    if(!user) return;
    const userUpper = user.toUpperCase();
    const equipeRenata = ['RENATA', 'HOZANA', 'ISRAEL', 'ROSANGELA', 'SARA', 'VINICIUS'];
    const equipeCarol  = ['CAROL', 'ALICE', 'CHARLENE', 'HEMILLY', 'MICHELLE'];
    
    window.OneSignalDeferred.push(async function(OneSignal) {
        if (equipeRenata.includes(userUpper)) {
            OneSignal.User.addTag("equipe", "equipe_renata");
        } else if (equipeCarol.includes(userUpper)) {
            OneSignal.User.addTag("equipe", "equipe_carol");
        }
    });
}

async function checkPushSubscription() {
    window.OneSignalDeferred = window.OneSignalDeferred || [];
    window.OneSignalDeferred.push(async function(OneSignal) {
        if (OneSignal.Notifications.permission === true) {
            const pushButton = document.getElementById('btnHabilitarPush');
            const pushStatus = document.getElementById('pushStatus');
            if (pushButton && pushStatus) {
                pushStatus.innerText = 'Você já está recebendo notificações.';
                pushButton.innerHTML = '<i class="fas fa-check-circle"></i> Inscrito';
                pushButton.disabled = true;
            }
        }
    });
}

async function enviarNotificacaoPush() {
    const target = document.getElementById('pushTarget').value;
    const title = document.getElementById('pushTitle').value;
    const body = document.getElementById('pushBody').value;
    const btn = document.querySelector('#painelAdminPush button');

    if (!body) {
        Swal.fire('Atenção', 'A mensagem não pode estar vazia.', 'warning');
        return;
    }

    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Enviando...';

    try {
        const response = await fetch(URL_PUSH_BACKEND, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                acao: "sendMessage",
                messagePayload: { target, title, body, url: "/" }
            })
        });
        const result = await response.text();
        Swal.fire('Sucesso!', result, 'success');
        document.getElementById('pushBody').value = '';
        document.getElementById('pushTitle').value = '';

    } catch (error) {
        Swal.fire('Erro!', 'Falha ao enviar a notificação. Verifique a conexão.', 'error');
        console.error("Erro ao enviar notificação:", error);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-paper-plane"></i> Enviar Mensagem';
    }
}

// --- SISTEMA DE CAIXA DE MENSAGENS (IN-APP) ---

function salvarNotificacaoLocal(payload) {
    let notifs = JSON.parse(localStorage.getItem('appNotificacoes') || '[]');
    notifs.unshift({
        title: payload.title || "Aviso Miss Rôse",
        body: payload.body || "Você tem uma nova mensagem.",
        date: new Date().toLocaleString('pt-BR'),
        lida: false
    });
    localStorage.setItem('appNotificacoes', JSON.stringify(notifs));
    atualizarBadgeNotificacoes();
}

function atualizarBadgeNotificacoes() {
    let notifs = JSON.parse(localStorage.getItem('appNotificacoes') || '[]');
    let naoLidas = notifs.filter(n => !n.lida).length;
    const badge = document.getElementById('badgeNotificacao');
    
    if (badge) {
        if (naoLidas > 0) {
            badge.innerText = naoLidas > 9 ? "9+" : naoLidas;
            badge.style.display = 'block';
        } else {
            badge.style.display = 'none';
        }
    }
}

function abrirPainelNotificacoes() {
    document.getElementById('modalNotificacoes').classList.remove('hidden');
    renderizarNotificacoes();
}

function fecharPainelNotificacoes() {
    document.getElementById('modalNotificacoes').classList.add('hidden');
    // Marca todas como lidas ao fechar o painel
    let notifs = JSON.parse(localStorage.getItem('appNotificacoes') || '[]');
    notifs.forEach(n => n.lida = true);
    localStorage.setItem('appNotificacoes', JSON.stringify(notifs));
    atualizarBadgeNotificacoes();
}

function renderizarNotificacoes() {
    let notifs = JSON.parse(localStorage.getItem('appNotificacoes') || '[]');
    const container = document.getElementById('listaNotificacoesPainel');
    if (notifs.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #888; margin-top: 20px;">Nenhuma mensagem no momento.</p>';
        return;
    }
    container.innerHTML = notifs.map(n => `
        <div style="background: ${n.lida ? 'transparent' : 'rgba(216, 27, 96, 0.05)'}; border: 1px solid ${n.lida ? '#ddd' : 'var(--primary)'}; border-left: 4px solid ${n.lida ? '#ccc' : 'var(--primary)'}; padding: 12px; border-radius: 6px; font-size: 13px; transition: 0.3s;">
            <strong style="display: block; color: var(--primary); margin-bottom: 5px; font-size: 14px;">${n.title}</strong>
            <p style="margin: 0 0 5px 0; color: var(--text);">${n.body}</p>
            <small style="color: #888;"><i class="far fa-clock"></i> Recebido em: ${n.date}</small>
        </div>
    `).join('');
}

function limparNotificacoes() {
    localStorage.setItem('appNotificacoes', '[]');
    renderizarNotificacoes();
    atualizarBadgeNotificacoes();
    Toast.fire({ icon: 'success', title: 'Caixa de entrada limpa!' });
}