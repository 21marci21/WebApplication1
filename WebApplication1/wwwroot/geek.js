var viccek;

function letöltés() {
    fetch('/jokes.json')
        .then(response => response.json())
        .then(data => letöltésBefejeződött(data))
}


function letöltésBefejeződött(d) {
    console.log("Sikeres letöltés")
    console.log(d)
    viccek = d;
    var geek = document.getElementById("geek")
    for (var i = 0; i < viccek.length; i++) {
        var újelem = document.createElement("div")
        újelem.innerHTML = viccek[i].question
        geek.appendChild(újelem)
    }
}
window.onload = letöltés();
console.log(viccek)