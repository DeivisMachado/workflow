(function () {
	const linhas = Array.from(
		document.querySelectorAll("body > table:nth-child(2) > tbody > tr")
	).slice(1);

	const valores = linhas
		.map((tr) => {
			const td = tr.querySelector("td:nth-child(24)");
			if (!td) return "";

			const texto = td.textContent.trim();
			if (/total/i.test(texto)) return "";

			return texto;
		})
		.filter((v) => v !== "");

	const lista = valores.join(",");
	const ta = document.createElement("textarea");
	ta.value = lista;
	document.body.appendChild(ta);
	ta.select();
	document.execCommand("copy");
	document.body.removeChild(ta);

	const msg = document.createElement("div");
	msg.innerText =
		valores.length + " ordens foram copiadas para a área de transferência";
	msg.style.position = "fixed";
	msg.style.bottom = "20px";
	msg.style.right = "20px";
	msg.style.background = "#1f2937";
	msg.style.color = "#fff";
	msg.style.padding = "12px 18px";
	msg.style.borderRadius = "8px";
	msg.style.boxShadow = "0 4px 10px rgba(0,0,0,0.3)";
	msg.style.fontFamily = "Arial, sans-serif";
	msg.style.fontSize = "14px";
	msg.style.zIndex = "9999";
	msg.style.opacity = "0";
	msg.style.transition = "opacity 0.3s ease";

	document.body.appendChild(msg);

	setTimeout(() => {
		msg.style.opacity = "1";
	}, 10);

	setTimeout(() => {
		msg.style.opacity = "0";
		setTimeout(() => {
			document.body.removeChild(msg);
		}, 300);
	}, 3000);
})();