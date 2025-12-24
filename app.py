async function doGenerate(e){
  e.preventDefault();

  const sheet = document.getElementById('sheet').value;
  const start = document.getElementById('start').value;
  const end   = document.getElementById('end').value;
  const mode  = document.querySelector('input[name="mode"]:checked').value;

  if(!sheet || !start || !end){
    alert("シート・開始・終了番号が必要です。");
    return false;
  }

  const win = window.open("about:blank", "_blank");

  const url = (mode === "pdf")
    ? "/generate"
    : "/generate_html_test";

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({sheet, start, end})
    });

    if(!res.ok){
      const tx = await res.text();
      win.close();
      alert("エラー: " + tx);
      return false;
    }

    if(mode === "pdf"){
      const data = await res.json();
      win.location.href = data.pdf_url;
    }else{
      const html = await res.text();
      win.document.open();
      win.document.write(html);
      win.document.close();
    }

  } catch(err){
    win.close();
    alert("通信エラー: " + err);
  }

  return false;
}
