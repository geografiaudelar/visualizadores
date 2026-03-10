
function explorar(tipo){

if(tipo === "interno"){

alert("Abrir visualizador de tesis");

}

if(tipo === "casos"){

alert("Abrir visualizador de casos de uso");

}

if(tipo === "proyectos"){

alert("Abrir visualizador de proyectos");

}

document.querySelectorAll('.btn-link.disabled').forEach(function(link){
link.addEventListener('click', function(e){
e.preventDefault();
});
});

document.querySelectorAll('.menu-disabled').forEach(function(link){
link.addEventListener('click', function(e){
e.preventDefault();
});
});

}