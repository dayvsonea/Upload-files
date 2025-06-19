import os, io, re, tempfile, subprocess, datetime, uuid
from flask import Flask, render_template, request, send_from_directory, redirect, url_for
from pptx import Presentation
import pandas as pd, requests

TEMPLATE_DIR=os.getenv("PPT_TEMPLATE_DIR","./templates_ppt")
WORKBOOK_XLSX=os.getenv("WORKBOOK_XLSX","")
UPLOAD_LIMIT=int(os.getenv("UPLOAD_LIMIT","5"))

OPTION_MAP={
 "1":"0000 1 Nova Venda COM Arquivo.pptx",
 "2":"0000 2 Nova Venda SEM Arquivo.pptx",
 "3":"0000 3 Upgrade COM Arquivo.pptx",
 "4":"0000 4 Upgrade SEM Arquivo.pptx"
}
REGEX=re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")

app=Flask(__name__)
app.secret_key="ea-secret"

def load_workbook():
    if not WORKBOOK_XLSX: return pd.DataFrame(), pd.DataFrame()
    data=requests.get(WORKBOOK_XLSX, timeout=20).content
    xls=pd.ExcelFile(io.BytesIO(data))
    planos=pd.read_excel(xls,"Planos")
    variaveis=pd.read_excel(xls,"Variáveis")
    return planos, variaveis

planos_df, variaveis_df = load_workbook()
down_labels = variaveis_df[variaveis_df["Variável"].str.startswith("down")]["Qual conteúdo"].tolist() if not variaveis_df.empty else [f"Arquivo {i}" for i in range(1,UPLOAD_LIMIT+1)]

def substitute_ppt(tpl, mp, out):
    prs=Presentation(tpl)
    def dist(p,t):
        idx=0
        for r in p.runs:
            ln=len(r.text); r.text=t[idx:idx+ln]; idx+=ln
        if idx<len(t): p.runs[-1].text+=t[idx:]
    def proc(p,sh):
        tx="".join(r.text for r in p.runs)
        if "{{" in tx:
            def rep(m):
                k=m.group(1).lower()
                return "Baixar Arquivo Aqui" if k.startswith("down") else mp.get(k,m.group(0))
            dist(p, REGEX.sub(rep, tx))
            dl=REGEX.fullmatch(tx.strip())
            if dl and dl.group(1).lower().startswith("down"):
                sh.click_action.hyperlink.address=mp[dl.group(1).lower()]
        if mp.get("valorsemdesc")==mp.get("valorcomdesc"):
            low=tx.lower()
            if "de:" in low: [setattr(r,"text","") for r in p.runs]
            elif "por:" in low: dist(p, mp["valorsemdesc"])
    def walk(s):
        if s.has_text_frame:
            for p in s.text_frame.paragraphs: proc(p,s)
        if s.has_table:
            for row in s.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs: proc(p,s)
        if s.shape_type==6:
            for sub in s.shapes: walk(sub)
    for slide in prs.slides:
        for sh in slide.shapes: walk(sh)
    prs.save(out)

def to_pdf(ppt):
    pdf=ppt.replace(".pptx",".pdf")
    try:
        subprocess.run(["soffice","--headless","--convert-to","pdf",ppt,"--outdir",os.path.dirname(ppt)],check=True,timeout=90)
        return pdf
    except Exception: return ""

@app.route("/")
def index(): return render_template("index.html")

@app.route("/form/<opt>")
def form(opt):
    if opt not in OPTION_MAP: return redirect(url_for("index"))
    return render_template("form.html", opt=opt, template_name=OPTION_MAP[opt], show_uploads=opt in ["1","3"], down_labels=down_labels)

@app.route("/generate/<opt>", methods=["POST"])
def generate(opt):
    if opt not in OPTION_MAP: return redirect(url_for("index"))
    mp={k.lower():v for k,v in request.form.items()}
    if 'desconto' not in request.form: mp["valorcomdesc"]=mp.get("valorsemdesc")
    row=planos_df[planos_df["Plano"].str.strip().str.lower()==mp.get("plano","").lower()]
    if not row.empty:
        r=row.iloc[0]
        mp.setdefault("descplan", r.get("Descrição",""))
        mp.setdefault("dimensionamento", r.get("Dimensionamento",""))
    mp["dataenvio"]=datetime.datetime.now().strftime("%d/%m/%Y")
    mp["hoje"]=mp["dataenvio"]
    for i in range(1,UPLOAD_LIMIT+1): mp[f"down{i}"]="https://www.e-auditoria.com.br"
    tmp=tempfile.mkdtemp()
    ppt_out=os.path.join(tmp,f"Proposta_{uuid.uuid4().hex}.pptx")
    substitute_ppt(os.path.join(TEMPLATE_DIR, OPTION_MAP[opt]), mp, ppt_out)
    pdf=to_pdf(ppt_out)
    email=render_template("email_template.txt", **mp)
    email_path=os.path.join(tmp,"mensagem_email.txt"); open(email_path,"w",encoding="utf-8").write(email)
    files=[("PPTX",os.path.basename(ppt_out))]
    if pdf: files.append(("PDF",os.path.basename(pdf)))
    files.append(("Mensagem (.txt)",os.path.basename(email_path)))
    return render_template("success.html", files=files, folder=os.path.basename(tmp))

@app.route("/download/<folder>/<filename>")
def download(folder,filename):
    return send_from_directory(os.path.join(tempfile.gettempdir(),folder), filename, as_attachment=True)

if __name__=="__main__":
    app.run(host="0.0.0.0", port=5000)
