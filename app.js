let modeResume = false;

let data =
JSON.parse(
localStorage.getItem("boutiqueData")
) || [];

function save(){

localStorage.setItem(
"boutiqueData",
JSON.stringify(data)
);

}

function toggleResume(){

modeResume = !modeResume;

render();

}

function render(){

const search =
document
.getElementById("search")
.value
.toLowerCase();
const titre =
document.getElementById("titreApp");

const toolbar =
document.getElementById("toolbarActions");

let html = "";
if(search.trim() !== ""){

    titre.style.display = "none";

    toolbar.querySelectorAll("button")
    .forEach(btn => {

        if(
            !btn.textContent.includes("Ajouter")
        ){
            btn.style.display = "none";
        }

    });

}
else{

    titre.style.display = "";

    toolbar.querySelectorAll("button")
    .forEach(btn => {

        btn.style.display = "";

    });

}
if(modeResume){

const articles =
data.filter(
a => Number(a.quantite) > 0
);

articles.forEach((item,index)=>{

html += `
<div class="card ${item.coche ? 'resumeOk' : ''}">

<label>

<input
type="checkbox"
${item.coche ? "checked" : ""}
onchange="toggleCheck(${index})">

${item.article}

</label>

<div>
Quantité : ${item.quantite}
</div>

</div>
`;

});

}
else{

data

.filter(item =>
item.article
.toLowerCase()
.includes(search)
)

.forEach((item,index)=>{

html += `
<div class="card">

<b>${item.article}</b>

<div>
Code : ${item.code}
</div>

<input
class="qtyInput"
id="qty_${item.code}"
type="number"
value="${
item.quantite === 0
? ""
: item.quantite
}">

<button
onclick="valider('${item.code}')"

OK
</button>

</div>
`;

});

}

document.getElementById("cards")
.innerHTML = html;

}

function valider(code){

    const article =
    data.find(
        a => String(a.code) === String(code)
    );

    if(!article) return;

    const input =
    document.getElementById(
        "qty_" + code
    );

    article.quantite =
    input.value === ""
    ? 0
    : parseInt(input.value) || 0;

    save();

    document.getElementById("search").value = "";

    render();

}

function toggleCheck(index){

data[index].coche =
!data[index].coche;

save();

render();

}

function addArticle(){

const code =
prompt("Code");

if(!code) return;

const article =
prompt("Article");

if(!article) return;

data.push({

code,
article,
quantite:0,
coche:false

});

save();

render();

}

function resetStock(){

if(
!confirm(
"Remettre toutes les quantités à zéro ?"
)
) return;

data.forEach(a=>{

a.quantite = 0;
a.coche = false;

});

save();

render();

}

function exportExcel(){

const wb =
XLSX.utils.book_new();

const ws =
XLSX.utils.json_to_sheet(data);

XLSX.utils.book_append_sheet(
wb,
ws,
"BOUTIQUE"
);

const d = new Date();

const fichier =
"boutique-" +
d.getFullYear() + "-" +
String(d.getMonth()+1).padStart(2,"0") + "-" +
String(d.getDate()).padStart(2,"0") + "-" +
String(d.getHours()).padStart(2,"0") + "h" +
String(d.getMinutes()).padStart(2,"0") +
".xlsx";

XLSX.writeFile(
wb,
fichier
);

}

function importExcel(event){

const file =
event.target.files[0];

if(!file) return;

const reader =
new FileReader();

reader.onload = function(e){

const workbook =
XLSX.read(
e.target.result,
{type:"array"}
);

const sheet =
workbook.Sheets[
workbook.SheetNames[0]
];

const rows =
XLSX.utils.sheet_to_json(sheet);

data = rows.map(r => ({

code:
r.NA ||
r.Code ||
r.code ||
"",

article:
r.ARTICLE ||
r.Article ||
r.article ||
"",

quantite:
Number(
r.QUANTITE ||
r.Quantite ||
r.quantite ||
0
),

coche:false

}));

save();

render();

alert(
"Import terminé : " +
data.length +
" articles"
);

};

reader.readAsArrayBuffer(file);

}

render();
