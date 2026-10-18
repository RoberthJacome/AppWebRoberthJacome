/* ============================================================
   PoliMusic - Reproductor global (JS puro)
   ============================================================ */
(function () {
    'use strict';

    var audio = document.getElementById('audio');
    var table = document.getElementById('songTable');
    if (!audio || !table) { return; }

    // Filas reproducibles (las que tienen archivo)
    var rows = Array.prototype.filter.call(
        table.querySelectorAll('tbody tr[data-src]'),
        function (row) { return row.getAttribute('data-src') !== ''; }
    );

    var el = {
        title:  document.getElementById('playerTitle'),
        artist: document.getElementById('playerArtist'),
        toggle: document.getElementById('btnToggle'),
        prev:   document.getElementById('btnPrev'),
        next:   document.getElementById('btnNext'),
        seek:   document.getElementById('seek'),
        cur:    document.getElementById('timeCurrent'),
        tot:    document.getElementById('timeTotal')
    };

    var index = -1;
    var arrastrando = false;

    // Segundos -> m:ss
    function formato(seg) {
        if (!isFinite(seg) || seg < 0) { seg = 0; }
        var m = Math.floor(seg / 60);
        var s = Math.floor(seg % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }

    // Estado visual de reproducción
    function estado(sonando) {
        document.body.classList.toggle('is-playing', sonando);
        el.toggle.setAttribute('aria-label', sonando ? 'Pausar' : 'Reproducir');
    }

    // Resaltar la fila activa
    function resaltar() {
        rows.forEach(function (row, i) {
            row.classList.toggle('active', i === index);
        });
    }

    // Cargar una canción de la lista
    function cargar(i, autoplay) {
        if (i < 0 || i >= rows.length) { return; }

        index = i;
        var row = rows[i];

        audio.src = row.getAttribute('data-src');
        el.title.textContent = row.getAttribute('data-title');
        el.artist.textContent = row.getAttribute('data-artist') || 'Artista desconocido';
        el.seek.value = 0;
        el.cur.textContent = '0:00';
        el.tot.textContent = '0:00';

        resaltar();
        if (autoplay) { reproducir(); }
    }

    function reproducir() {
        var p = audio.play();
        if (p && typeof p.catch === 'function') {
            p.catch(function () { estado(false); });
        }
    }

    // Play / pausa
    function alternar() {
        if (index < 0) {
            if (rows.length) { cargar(0, true); }
            return;
        }
        if (audio.paused) { reproducir(); } else { audio.pause(); }
    }

    /* ---------- Eventos del <audio> ---------- */
    audio.addEventListener('play',  function () { estado(true); });
    audio.addEventListener('pause', function () { estado(false); });

    // Al terminar, pasar a la siguiente
    audio.addEventListener('ended', function () {
        if (index < rows.length - 1) {
            cargar(index + 1, true);
        } else {
            estado(false);
        }
    });

    audio.addEventListener('loadedmetadata', function () {
        el.tot.textContent = formato(audio.duration);
    });

    audio.addEventListener('timeupdate', function () {
        if (arrastrando || !audio.duration) { return; }
        el.seek.value = (audio.currentTime / audio.duration) * 100;
        el.cur.textContent = formato(audio.currentTime);
    });

    audio.addEventListener('error', function () {
        if (!audio.getAttribute('src')) { return; }
        el.artist.textContent = 'No se pudo cargar el archivo de audio';
        estado(false);
    });

    /* ---------- Barra de progreso ---------- */
    el.seek.addEventListener('input', function () {
        arrastrando = true;
        if (audio.duration) {
            el.cur.textContent = formato((el.seek.value / 100) * audio.duration);
        }
    });

    el.seek.addEventListener('change', function () {
        if (audio.duration) {
            audio.currentTime = (el.seek.value / 100) * audio.duration;
        }
        arrastrando = false;
    });

    /* ---------- Controles ---------- */
    el.toggle.addEventListener('click', alternar);

    el.prev.addEventListener('click', function () {
        if (!rows.length) { return; }
        cargar(index <= 0 ? 0 : index - 1, true);
    });

    el.next.addEventListener('click', function () {
        if (!rows.length) { return; }
        cargar(index < 0 || index >= rows.length - 1 ? 0 : index + 1, true);
    });

    /* ---------- Filas de la tabla ---------- */
    rows.forEach(function (row, i) {
        row.addEventListener('click', function (e) {
            // Los enlaces de Editar / Eliminar siguen su curso
            if (e.target.closest('a')) { return; }

            if (i === index) {
                alternar();
            } else {
                cargar(i, true);
            }
        });
    });

    estado(false);
}());
