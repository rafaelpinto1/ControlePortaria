const msalConfig = {
  auth: {
    clientId: "2123d101-f699-4ee5-a09a-3a3882618ec9",
    authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
    redirectUri: window.location.href,
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function loginEObterToken() {
  const loginResponse = await msalInstance.loginPopup({
    scopes: ["User.Read", "Sites.ReadWrite.All", "Mail.Send"],
  });

  const tokenResponse = await msalInstance.acquireTokenSilent({
    scopes: ["Sites.ReadWrite.All", "Mail.Send"],
    account: loginResponse.account,
  });

  return tokenResponse.accessToken;
}

function calcularDuracao(entrada, saida) {
  const [h1, m1] = entrada.split(":").map(Number);
  const [h2, m2] = saida.split(":").map(Number);
  return ((h2 * 60 + m2) - (h1 * 60 + m1)) / 60;
}

async function enviarDados(e) {
  e.preventDefault();

  const form = document.getElementById("form-portaria");
  const dados = {
    Title: form.nome.value.trim(),
    Data: form.data.value,
    TipoServico: form.tipoServico.value.trim(),
    Setor: form.setor.value.trim(),
    VeiculoPlaca: form.veiculoPlaca.value.trim(),
    Destino: form.destino.value.trim(),
    Entrada: form.entrada.value,
    Saida: form.saida.value,
    Observacoes: form.observacoes.value.trim(),
  };

  const duracao = calcularDuracao(dados.Entrada, dados.Saida);
  const token = await loginEObterToken();

  const siteUrl = "https://gsilvainfo.sharepoint.com/sites/Inf";
  const listName = "ControlePortaria";

  try {
    const response = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify({
          __metadata: { type: "SP.Data.ControlePortariaListItem" },
          ...dados,
        }),
      }
    );

    if (!response.ok) throw new Error("Erro ao gravar no SharePoint");

    const ehGSilva = dados.Setor.toLowerCase().includes("g. silva");
    if (ehGSilva && (duracao < 7 || duracao > 9)) {
      await fetch("https://<NOME_DA_FUNCTION>.azurewebsites.net/api/notificarHorario", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...dados, Duracao: duracao }),
      });
    }

    alert("Registro enviado com sucesso!");
    form.reset();
  } catch (err) {
    console.error("Erro ao enviar:", err);
    alert("Erro ao enviar os dados.");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("form-portaria").addEventListener("submit", enviarDados);
});
