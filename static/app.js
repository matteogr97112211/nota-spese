async function uploadOCR(input){
const file=input.files[0];
const formData=new FormData();
formData.append("file",file);

const res=await fetch("/ocr",{method:"POST",body:formData});
const data=await res.json();

document.getElementById("importo").value=data.importo;
document.getElementById("data").value=data.data;

console.log(data.raw);
}

function genera(){
fetch("/generate",{method:"POST"})
.then(res=>res.blob())
.then(blob=>{
const url=window.URL.createObjectURL(blob);
const a=document.createElement("a");
a.href=url;
a.download="nota_spese.xlsx";
a.click();
});
}