let sistemaSelecionado = null;

// Define o sistema escolhido
function selecionarSistema(nome) {
  sistemaSelecionado = nome;
  console.log("[LOG] Sistema selecionado:", nome);
}

function exibirErro(mensagem) {
  const alerta = document.getElementById("alerta");
  alerta.textContent = mensagem;
  alerta.classList.remove("d-none");
  alerta.classList.add("show");

  // Oculta automaticamente após 6 segundos
  setTimeout(() => {
    alerta.classList.add("d-none");
    alerta.classList.remove("show");
    alerta.textContent = "";
  }, 6000);
}

// Inicia o processo de geração e download da planilha
async function buscar() {
  const botao = document.getElementById("btnBuscar");
  const loading = document.getElementById("loading");

  if (!sistemaSelecionado) {
    alert("Selecione um sistema primeiro.");
    return;
  }

  const dataInput = document.getElementById("data");
  const data = dataInput.value;

  if (!data) {
    alert("Informe a data para buscar os contratos.");
    return;
  }

  const url = `http://localhost:3000/api/relatorio?sistema=${sistemaSelecionado}&data=${data}`;
  console.log("[LOG] Iniciando requisição para:", url);

  botao.disabled = true;
  loading.style.display = "block";

  try {
    const response = await fetch(url);

    if (response.status === 429) {
      const erro = await response.json();
      alert(erro.erro || "Já existe uma requisição em andamento. Aguarde.");
      return;
    }

    if (!response.ok) {
      let erroMsg = "Erro ao gerar planilha.";
      try {
        const erro = await response.json();
        erroMsg = erro.erro || erroMsg;
      } catch {
        const texto = await response.text();
        console.error("[ERRO] Conteúdo inesperado:", texto);
      }

      alert(erroMsg);
      return;
    }

    console.log("[LOG] Download da planilha iniciado.");
    const blob = await response.blob();
    const downloadUrl = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = `relatorio_${data}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(downloadUrl);

    console.log("[LOG] Download concluído com sucesso.");
  } catch (e) {
    console.error("[FALHA] Erro ao tentar baixar a planilha:", e);
    alert("Erro ao baixar a planilha.");
  } finally {
    botao.disabled = false;
    loading.style.display = "none";
  }
}
