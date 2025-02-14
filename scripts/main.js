// filepath: /github-pages-website/github-pages-website/scripts/main.js
document.addEventListener("DOMContentLoaded", function() {
    document.getElementById("generateDoc").addEventListener("click", generateDoc);
});

async function generateDoc() {
    const fileInput = document.getElementById("file");
    if (!fileInput.files.length) {
        alert("Selecione um arquivo DOCX primeiro.");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = async function(event) {
        const content = event.target.result;
        const zip = new JSZip();
        await zip.loadAsync(content);
        
        const docText = await zip.file("word/document.xml").async("string");

        const getValue = (id) => document.getElementById(id).value || "";

        let updatedText = docText
            .replace(/Data:/g, `Data: ${getValue("data")}`)
            .replace(/Responsável pela visita:/g, `Responsável pela visita: ${getValue("responsavel")}`)
            .replace(/Horário:/g, `Horário: ${getValue("horario")}`)
            .replace(/Tipo de Reunião:/g, `Tipo de Reunião: ${getValue("tipo")}`)
            .replace(/Local:/g, `Local: ${getValue("local")}`)
            .replace(/Temas:/g, `Temas: ${getValue("temas")}`)
            .replace(/Participantes:/g, `Participantes: ${getValue("participantes")}`)
            .replace(/Assuntos tratados:/g, `Assuntos tratados: ${getValue("assuntos")}`)
            .replace(/Recomendações:/g, `Recomendações: ${getValue("recomendacoes")}`)
            .replace(/Conclusões:/g, `Conclusões: ${getValue("conclusoes")}`);

        zip.file("word/document.xml", updatedText);

        zip.generateAsync({ type: "blob" }).then(function(blob) {
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "ATA_Preenchida.docx";
            link.click();
        });
    };
    
    reader.readAsArrayBuffer(file);
}