window.onload = () => {
    console.log("betöltődött")
}
var faktor = n => {
    if (n == 0 || n == 1) {
        return 1;
    }
    else {
        return n * faktor(n - 1)
    }
}

function pascalharomszog(sor, oszlop) {
    szam = faktor(sor) / (faktor(oszlop) * faktor(sor - oszlop))
    return szam
}

var pascal = document.getElementById("pascal")
var méret = 10
for (var s = 0; s < méret; s++) {
    var újsor = document.createElement("div")
    újsor.classList.add("sor")
    //újsor.innerHTML = faktor(s)
    pascal.appendChild(újsor)

    for (var o = 0; o <= s; o++) {
        var újelem = document.createElement("div")
        újelem.innerHTML = pascalharomszog(s, o);
        //újelem.innerHTML = `${s}:${o}`;
        újelem.classList.add("elem")
        újsor.appendChild(újelem)
    }
}

