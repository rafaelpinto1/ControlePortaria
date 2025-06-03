document.addEventListener("DOMContentLoaded", () => {
  const msalConfig = {
    auth: {
      clientId: "2123d101-f699-4ee5-a09a-3a3882618ec9",
      authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
      redirectUri: window.location.href,
    },
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const btnLogin = document.getElementById("btn-login");
  const form = document.getElementById("form-portaria");
  let currentAccount = null;
  let accessToken = "";

  btnLogin.addEventListener("click", async () => {
    try {
      const loginResponse = await msalInstance.loginPopup({
        scopes: ["User.Read", "Sites.ReadWrite.All", "Mail.Send"],
      });
      currentAccount = loginResponse.account;
      msalInstance.setActiveAccount(currentAccount);

      // Obter token
      const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.ReadWrite.All", "Mail.Send"],
        account: currentAccount,
      });
      accessToken = tokenResponse.accessToken;

      btnLogin.style.display = "none";
      form.style.display = "block";
      document.getElementById("status").innerText = "Logado como: " + currentAccount.username;
    } catch (err) {
      alert("Falha no login: " + err.message);
    }
  });

  function calcularDuracao(entrada, saida) {
    const [h1, m1] = entrada.split(":").map(Number);
    const [h2, m2] = saida.split(":").map(Number);
    return ((h2 * 60 + m2) - (h1 * 60 + m1)) / 60;
  }

  async function enviarDados(e) {
    e.preventDefault();

    if (!accessToken) {
      alert("Você precisa fazer login antes de enviar.");
      return;
    }

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

    try {
      // Obter siteId do SharePoint
      const siteResp = await fetch(
        "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Inf",
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );

      if (!siteResp.ok) throw new Error("Erro ao obter site do SharePoint");

      const siteJson = await siteResp.json();
      const siteId = siteJson.id.split(",")[1]; // Pega só o GUID do site

      // Enviar item para lista ControlePortaria via Graph API
      const postResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/ControlePortaria/items`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ fields: dados }),
        }
      );

      if (!postResp.ok) {
        const errorJson = await postResp.json();
        throw new Error(errorJson.error.message);
      }

      // Se setor é "g. silva" e duração < 7 ou > 9, chama Azure Function
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
      alert("Erro ao enviar os dados: " + err.message);
    }
  }

  form.addEventListener("submit", enviarDados);
});
