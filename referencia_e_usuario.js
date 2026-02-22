(() => {
  const tabela = document.querySelector("body > table:nth-child(2)");
  if (!tabela) {
    alert("Tabela não encontrada.");
    return;
  }

  const linhas = [...tabela.querySelectorAll("tbody > tr")].slice(1); // descarta a 1ª linha
  const dados = linhas
    .map((tr) => {
      const tds = tr.querySelectorAll("td");
      const referencia = tds[14]?.textContent?.trim(); // coluna 15
      const usuario = tds[32]?.textContent?.trim();    // coluna 33
      return referencia && usuario ? { referencia, usuario } : null;
    })
    .filter(Boolean);

  if (!dados.length) {
    alert("Nenhum dado válido encontrado (colunas 15 e 33).");
    return;
  }

  const resumoMap = dados.reduce((acc, item) => {
    acc[item.usuario] = (acc[item.usuario] || 0) + 1;
    return acc;
  }, {});

  const resumo = Object.entries(resumoMap)
    .map(([usuario, quantidade]) => ({ usuario, quantidade }))
    .sort((a, b) => b.quantidade - a.quantidade || a.usuario.localeCompare(b.usuario));

  document.getElementById("modal-ref-user")?.remove();

  const overlay = document.createElement("div");
  overlay.id = "modal-ref-user";
  overlay.style.cssText = `
    position: fixed; inset: 0; background: rgba(0,0,0,.45);
    display: flex; align-items: center; justify-content: center;
    z-index: 99999; font-family: Arial, sans-serif;
  `;

  const modal = document.createElement("div");
  modal.style.cssText = `
    background: #fff; width: min(900px, 95vw); max-height: 85vh;
    border-radius: 10px; box-shadow: 0 10px 30px rgba(0,0,0,.25);
    padding: 16px; overflow: auto;
  `;

  const gerarCsv = () => {
    const escapeCsv = (valor) => {
      const texto = String(valor ?? "").replace(/"/g, '""');
      return `"${texto}"`;
    };

    const linhasCsv = [];
    linhasCsv.push("RESUMO");
    linhasCsv.push(["Usuario", "Quantidade de referencias"].map(escapeCsv).join(";"));
    resumo.forEach((r) => {
      linhasCsv.push([r.usuario, r.quantidade].map(escapeCsv).join(";"));
    });

    linhasCsv.push("");
    linhasCsv.push("DETALHES");
    linhasCsv.push(["Referencia", "Usuario"].map(escapeCsv).join(";"));
    dados.forEach((d) => {
      linhasCsv.push([d.referencia, d.usuario].map(escapeCsv).join(";"));
    });

    return "\uFEFF" + linhasCsv.join("\r\n");
  };

  const baixarCsv = () => {
    const blob = new Blob([gerarCsv()], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const agora = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const sufixo = `${agora.getFullYear()}${pad(agora.getMonth() + 1)}${pad(agora.getDate())}_${pad(agora.getHours())}${pad(agora.getMinutes())}${pad(agora.getSeconds())}`;
    a.href = url;
    a.download = `resumo_detalhes_referencias_${sufixo}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const carregarSheetJs = () => {
    return new Promise((resolve, reject) => {
      if (window.XLSX) {
        resolve(window.XLSX);
        return;
      }

      const scriptExistente = document.getElementById("sheetjs-xlsx-cdn");
      if (scriptExistente) {
        scriptExistente.addEventListener("load", () => resolve(window.XLSX));
        scriptExistente.addEventListener("error", () => reject(new Error("Falha ao carregar SheetJS.")));
        return;
      }

      const script = document.createElement("script");
      script.id = "sheetjs-xlsx-cdn";
      script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      script.async = true;
      script.onload = () => {
        if (window.XLSX) {
          resolve(window.XLSX);
          return;
        }
        reject(new Error("SheetJS carregado, mas XLSX não está disponível."));
      };
      script.onerror = () => reject(new Error("Falha ao carregar SheetJS."));
      document.head.appendChild(script);
    });
  };

  const baixarXlsx = async () => {
    try {
      await carregarSheetJs();

      const resumoAba = [
        ["Usuario", "Quantidade de referencias"],
        ...resumo.map((r) => [r.usuario, r.quantidade]),
      ];

      const detalhesAba = [
        ["Referencia", "Usuario"],
        ...dados.map((d) => [d.referencia, d.usuario]),
      ];

      const workbook = window.XLSX.utils.book_new();
      const wsResumo = window.XLSX.utils.aoa_to_sheet(resumoAba);
      const wsDetalhes = window.XLSX.utils.aoa_to_sheet(detalhesAba);

      window.XLSX.utils.book_append_sheet(workbook, wsResumo, "Resumo");
      window.XLSX.utils.book_append_sheet(workbook, wsDetalhes, "Detalhes");

      const conteudoXlsx = window.XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
      });

      const blob = new Blob([conteudoXlsx], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const agora = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      const sufixo = `${agora.getFullYear()}${pad(agora.getMonth() + 1)}${pad(agora.getDate())}_${pad(agora.getHours())}${pad(agora.getMinutes())}${pad(agora.getSeconds())}`;

      a.href = url;
      a.download = `resumo_detalhes_referencias_${sufixo}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (erro) {
      console.error("Erro ao gerar XLSX:", erro);
      alert("Não foi possível gerar XLSX neste ambiente. Será baixado CSV.");
      baixarCsv();
    }
  };

  modal.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <h3 id="titulo-modal-ref-user" style="margin:0">Resumo por Usuário (${resumo.length})</h3>
      <div style="display:flex;gap:8px;align-items:center">
        <button id="detalhes-modal-ref-user" style="border:0;background:#1565c0;color:#fff;padding:8px 12px;border-radius:6px;cursor:pointer">Detalhes</button>
        <button id="baixar-modal-ref-user" style="border:0;background:#2e7d32;color:#fff;padding:8px 12px;border-radius:6px;cursor:pointer">Baixar</button>
        <button id="fechar-modal-ref-user" style="border:0;background:#e53935;color:#fff;padding:8px 12px;border-radius:6px;cursor:pointer">Fechar</button>
      </div>
    </div>
    <div id="conteudo-modal-ref-user"></div>
  `;

  const style = document.createElement("style");
  style.textContent = `
    #modal-ref-user .linha-tabela-modal-ref-user td {
      transition: background-color .12s ease;
    }
    #modal-ref-user .linha-tabela-modal-ref-user:hover td {
      background-color: #eef5ff !important;
    }
  `;
  modal.appendChild(style);

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  const titulo = document.getElementById("titulo-modal-ref-user");
  const conteudo = document.getElementById("conteudo-modal-ref-user");
  const botaoDetalhes = document.getElementById("detalhes-modal-ref-user");
  const botaoBaixar = document.getElementById("baixar-modal-ref-user");

  let exibindoDetalhes = false;
  let usuarioSelecionado = null;

  const renderizar = () => {
    if (!conteudo || !titulo || !botaoDetalhes) return;

    if (exibindoDetalhes) {
      const dadosFiltrados = usuarioSelecionado
        ? dados.filter((d) => d.usuario === usuarioSelecionado)
        : dados;

      const detalhesFiltradosHtml = dadosFiltrados
        .map(
          (d, i) => `
            <tr class="linha-tabela-modal-ref-user">
              <td style="padding:8px;border:1px solid #ddd;text-align:center">${i + 1}</td>
              <td style="padding:8px;border:1px solid #ddd">${d.referencia}</td>
              <td style="padding:8px;border:1px solid #ddd">${d.usuario}</td>
            </tr>
          `
        )
        .join("");

      titulo.textContent = usuarioSelecionado
        ? `Detalhes de ${usuarioSelecionado} (${dadosFiltrados.length})`
        : `Detalhes Referência x Usuário (${dadosFiltrados.length})`;
      botaoDetalhes.textContent = "Resumo";
      conteudo.innerHTML = `
        <table style="width:100%;border-collapse:collapse;font-size:14px">
          <thead>
            <tr style="background:#f5f5f5">
              <th style="padding:8px;border:1px solid #ddd">#</th>
              <th style="padding:8px;border:1px solid #ddd">Referência</th>
              <th style="padding:8px;border:1px solid #ddd">Usuário</th>
            </tr>
          </thead>
          <tbody>${detalhesFiltradosHtml}</tbody>
        </table>
      `;
      return;
    }

    titulo.textContent = `Resumo por Usuário (${resumo.length})`;
    botaoDetalhes.textContent = "Detalhes";
    conteudo.innerHTML = `
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <thead>
          <tr style="background:#f5f5f5">
            <th style="padding:8px;border:1px solid #ddd">#</th>
            <th style="padding:8px;border:1px solid #ddd">Usuário</th>
            <th style="padding:8px;border:1px solid #ddd">Quantidade de Referências</th>
          </tr>
        </thead>
        <tbody>
          ${resumo
            .map(
              (r, i) => `
                <tr class="linha-resumo-modal-ref-user linha-tabela-modal-ref-user" data-usuario="${r.usuario.replace(/"/g, "&quot;")}" style="cursor:pointer">
                  <td style="padding:8px;border:1px solid #ddd;text-align:center">${i + 1}</td>
                  <td style="padding:8px;border:1px solid #ddd">${r.usuario}</td>
                  <td style="padding:8px;border:1px solid #ddd;text-align:center">${r.quantidade}</td>
                </tr>
              `
            )
            .join("")}
        </tbody>
      </table>
      <div style="margin-top:8px;font-size:12px;color:#555">Clique em um usuário para ver apenas as referências dele.</div>
    `;

    conteudo.querySelectorAll(".linha-resumo-modal-ref-user").forEach((linha) => {
      linha.addEventListener("click", () => {
        usuarioSelecionado = linha.getAttribute("data-usuario");
        exibindoDetalhes = true;
        renderizar();
      });
    });
  };

  renderizar();

  botaoDetalhes.onclick = () => {
    exibindoDetalhes = !exibindoDetalhes;
    if (!exibindoDetalhes) {
      usuarioSelecionado = null;
    }
    renderizar();
  };

  botaoBaixar.onclick = () => {
    baixarXlsx();
  };

  document.getElementById("fechar-modal-ref-user").onclick = () => overlay.remove();
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });
})();