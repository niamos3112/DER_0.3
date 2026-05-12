// ==========================================
// CONFIGURACIÓN GLOBAL (PRODUCCIÓN EN FLASK)
// ==========================================
const API_URL = '/api/data';

let appData = {
    linkPostulacion: "#",
    directorio: [],
    eventos: [],
    anexos: []
};

let isAdmin = false;
let fechaActual = new Date();

// ==========================================
// 1. COMUNICACIÓN CON EL SERVIDOR (PYTHON)
// ==========================================
async function cargarDatosServidor() {
    try {
        const respuesta = await fetch(API_URL);
        if (respuesta.ok) {
            appData = await respuesta.json();
            
            // Ocultar banner de carga
            const banner = document.getElementById('loading-banner');
            if(banner) {
                banner.style.opacity = '0';
                setTimeout(() => banner.style.display = 'none', 500);
            }
            
            renderizarTodo();
        }
    } catch (error) {
        console.error("Error conectando al servidor Python:", error);
        const banner = document.getElementById('loading-banner');
        if(banner) {
            banner.innerText = "Error de conexión. Asegúrate de que el servidor esté activo.";
            banner.style.background = "#e31837";
        }
    }
}

async function guardarDatosServidor() {
    try {
        await fetch(API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(appData)
        });
        renderizarTodo();
    } catch (error) {
        alert("Error al guardar en el servidor");
    }
}

function renderizarTodo() {
    // Renderizar enlace de postulación (solo si existe el botón en la página actual)
    const btnPostula = document.getElementById('btn-postula-publico');
    const inputLink = document.getElementById('admin-link-postula');
    
    if (btnPostula) btnPostula.href = appData.linkPostulacion;
    if (inputLink) inputLink.value = appData.linkPostulacion;
    
    renderizarEntidadesAdmin();
    renderizarCalendario();
    renderizarAnexos();
}

document.addEventListener('DOMContentLoaded', cargarDatosServidor);

// ==========================================
// 2. CONTROL DE ACCESO (ADMINISTRADOR)
// ==========================================
function abrirModal(id) { 
    const modal = document.getElementById(id);
    if(modal) modal.classList.remove('hidden'); 
}

function cerrarModal(id) { 
    const modal = document.getElementById(id);
    if(modal) modal.classList.add('hidden'); 
}

function verificarAcceso() {
    if(document.getElementById('admin-pass').value === "admin123") {
        cerrarModal('admin-login');
        document.getElementById('admin-panel').classList.remove('hidden');
        document.getElementById('admin-pass').value = '';
        isAdmin = true;
        renderizarCalendario(); // Mostrar botones de edición
    } else { 
        alert("Contraseña incorrecta"); 
    }
}

function cerrarAdmin() { 
    document.getElementById('admin-panel').classList.add('hidden'); 
    isAdmin = false;
    renderizarCalendario();
}

function actualizarLinkPostulacion() {
    let url = document.getElementById('admin-link-postula').value;
    if (url && !url.startsWith('http')) url = 'https://' + url;
    appData.linkPostulacion = url;
    guardarDatosServidor();
    alert("Enlace actualizado para todos los usuarios.");
}

// ==========================================
// 3. GESTIÓN DE DIRECTORIOS (CON CATEGORÍAS)
// ==========================================
function agregarDirectorioManual() {
    const tipo = document.getElementById('dir-tipo').value;
    const categoria = document.getElementById('dir-categoria').value;
    const entidad = document.getElementById('dir-entidad').value;
    const nombre = document.getElementById('dir-nombre').value;
    
    if(!entidad || !nombre) return alert("La Entidad y el Nombre son obligatorios");

    appData.directorio.push({
        Tipo: tipo, 
        Categoria: categoria, // Guardamos la página a la que pertenece
        Entidad: entidad.trim(), 
        Integrante: nombre.trim(),
        Cargo: document.getElementById('dir-cargo').value.trim(),
        Celular: document.getElementById('dir-cel').value.trim(),
        Correo: document.getElementById('dir-correo').value.trim(),
        Observaciones: document.getElementById('dir-obs').value.trim()
    });

    guardarDatosServidor();
    
    // Limpiar campos de texto del panel
    document.querySelectorAll('.admin-sub-section input[type="text"]').forEach(el => el.value = '');
    alert(`Integrante añadido correctamente a la sección: ${categoria}`);
}

function descargarPlantilla() {
    // La plantilla Excel es simple. La categoría y el tipo se eligen en la página web antes de subir.
    const encabezados = [["Entidad", "Integrante", "Cargo", "Celular", "Correo", "Observaciones"]];
    const ws = XLSX.utils.aoa_to_sheet(encabezados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Datos");
    XLSX.writeFile(wb, "Formato_Directorio.xlsx");
}

function manejarSubidaExcel(evento) {
    const archivo = evento.target.files[0];
    const tipoSeleccionado = document.getElementById('dir-tipo').value;
    const categoriaSeleccionada = document.getElementById('dir-categoria').value;
    if (!archivo) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const datosJSON = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        // Mapeo seguro con sanitización (.trim) y asignación del Tipo y Categoría de la interfaz
        const datosMapeados = datosJSON.map(fila => ({
            Tipo: tipoSeleccionado,
            Categoria: categoriaSeleccionada,
            Entidad: (fila.Entidad || 'Entidad Desconocida').toString().trim(),
            Integrante: (fila.Integrante || '-').toString().trim(),
            Cargo: (fila.Cargo || '-').toString().trim(),
            Celular: (fila.Celular || '-').toString().trim(),
            Correo: (fila.Correo || '-').toString().trim(),
            Observaciones: (fila.Observaciones || '-').toString().trim()
        }));
        
        appData.directorio = [...appData.directorio, ...datosMapeados];
        guardarDatosServidor();
        document.getElementById('excel-input').value = ''; 
        alert(`Excel cargado correctamente como ${tipoSeleccionado} en ${categoriaSeleccionada}`);
    };
    reader.readAsArrayBuffer(archivo);
}

function eliminarEntidad(nombreEntidad) {
    if(confirm(`¿Seguro que deseas eliminar la entidad "${nombreEntidad}" y todos sus miembros?`)) {
        appData.directorio = appData.directorio.filter(d => d.Entidad !== nombreEntidad);
        guardarDatosServidor();
    }
}

function renderizarEntidadesAdmin() {
    const lista = document.getElementById('lista-admin-entidades');
    if(!lista) return;

    const entidades = [...new Set(appData.directorio.map(item => item.Entidad))];
    lista.innerHTML = ''; 
    
    entidades.forEach(ent => {
        // Obtenemos los detalles de la primera persona de esa entidad para mostrar la info en el panel
        const info = appData.directorio.find(d => d.Entidad === ent);
        const li = document.createElement('li');
        
        // Mostramos el nombre, si es Incubadora/Formulador, y en qué página está
        li.innerHTML = `<span><strong>${ent}</strong> <br><small style="color:#555;">(${info.Tipo} - ${info.Categoria})</small></span>`;
        
        const btnEliminar = document.createElement('button');
        btnEliminar.className = 'btn-danger';
        btnEliminar.innerHTML = '<i class="fa-solid fa-trash"></i>';
        btnEliminar.onclick = () => eliminarEntidad(ent);
        
        li.appendChild(btnEliminar);
        lista.appendChild(li);
    });
}

function abrirDirectorio(tipo, categoria) {
    document.getElementById('titulo-directorio').innerText = tipo === 'incubadora' ? 'Incubadoras y Otras Entidades' : 'Formuladores';
    
    // FILTRO DOBLE: Comprobamos el tipo y la página (categoría) a la que pertenece
    const data = appData.directorio.filter(d => d.Tipo === tipo && d.Categoria === categoria);
    
    const contenedor = document.getElementById('contenedor-entidades');
    contenedor.innerHTML = data.length === 0 ? '<p>No hay registros en esta sección para esta página.</p>' : '';

    const agrupado = data.reduce((acc, curr) => {
        if(!acc[curr.Entidad]) acc[curr.Entidad] = [];
        acc[curr.Entidad].push(curr);
        return acc;
    }, {});

    for (const [entidad, integrantes] of Object.entries(agrupado)) {
        let filas = integrantes.map(i => `<tr>
            <td data-label="Integrante">${i.Integrante||'-'}</td>
            <td data-label="Cargo">${i.Cargo||'-'}</td>
            <td data-label="Celular">${i.Celular||'-'}</td>
            <td data-label="Correo">${i.Correo||'-'}</td>
            <td data-label="Observaciones">${i.Observaciones||'-'}</td>
        </tr>`).join('');
        
        contenedor.innerHTML += `
            <div class="entidad-group">
                <div class="entidad-header"><i class="fa-solid fa-building"></i> ${entidad}</div>
                <div class="table-responsive">
                    <table>
                        <thead><tr><th>Integrante</th><th>Cargo</th><th>Celular</th><th>Correo</th><th>Observaciones</th></tr></thead>
                        <tbody>${filas}</tbody>
                    </table>
                </div>
            </div>`;
    }
    abrirModal('modal-directorio');
}

// ==========================================
// 4. GESTIÓN DEL CALENDARIO Y EVENTOS
// ==========================================
function cambiarMes(offset) {
    fechaActual.setMonth(fechaActual.getMonth() + offset);
    renderizarCalendario();
}

function renderizarCalendario() {
    const grid = document.getElementById('calendar-body');
    const displayMes = document.getElementById('mes-anio-display');
    const boxDetalles = document.getElementById('evento-detalles-box');
    
    if(!grid || !displayMes) return; // Por si estamos en una página sin calendario

    const mes = fechaActual.getMonth();
    const anio = fechaActual.getFullYear();
    const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    
    displayMes.innerText = `${meses[mes]} ${anio}`;
    grid.innerHTML = '';
    
    if(boxDetalles) boxDetalles.classList.add('hidden');
    
    const diasEnMes = new Date(anio, mes + 1, 0).getDate();

    for(let i=1; i<=diasEnMes; i++) {
        const dia = document.createElement('div');
        dia.className = 'calendar-day';
        dia.innerText = i;
        
        const fechaStr = `${anio}-${String(mes + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
        const eventosDelDia = appData.eventos.filter(e => e.fecha === fechaStr);
        
        if(eventosDelDia.length > 0) {
            dia.classList.add('has-event');
            dia.onclick = () => mostrarDetalleEvento(eventosDelDia);
        }
        grid.appendChild(dia);
    }
}

function mostrarDetalleEvento(eventosDelDia) {
    const box = document.getElementById('evento-detalles-box');
    if(!box) return;

    box.innerHTML = '';

    eventosDelDia.forEach(evento => {
        let ubicacionHTML = `<span><i class="fa-solid fa-location-dot"></i> ${evento.ubicacion}</span>`;
        
        if(evento.ubicacion.includes('http') || evento.ubicacion.includes('meet') || evento.ubicacion.includes('zoom')) {
            let urlSegura = evento.ubicacion.startsWith('http') ? evento.ubicacion : 'https://' + evento.ubicacion;
            ubicacionHTML = `<a href="${urlSegura}" target="_blank" rel="noopener noreferrer" class="btn-meet-link"><i class="fa-solid fa-video"></i> Ir a la Reunión / Ubicación</a>`;
        }

        let adminHTML = isAdmin ? `
            <div class="admin-event-actions">
                <button onclick="prepararEdicionEvento(${evento.id})" class="btn-secondary" style="width:50%"><i class="fa-solid fa-pen"></i> Editar</button>
                <button onclick="eliminarEvento(${evento.id})" class="btn-danger" style="width:50%"><i class="fa-solid fa-trash"></i> Eliminar</button>
            </div>
        ` : '';

        box.innerHTML += `
            <div style="margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom:10px;">
                <div class="evento-hora"><i class="fa-regular fa-clock"></i> ${evento.hora || 'Todo el día'}</div>
                <div class="evento-titulo">${evento.titulo}</div>
                ${ubicacionHTML}
                ${adminHTML}
            </div>
        `;
    });
    box.classList.remove('hidden');
}

function guardarEvento() {
    const id = document.getElementById('evento-edit-id').value;
    const fecha = document.getElementById('evento-fecha').value;
    const hora = document.getElementById('evento-hora').value;
    const titulo = document.getElementById('evento-titulo').value;
    const ubicacion = document.getElementById('evento-ubicacion').value;

    if(!fecha || !titulo) return alert("Fecha y título son obligatorios");

    if (id) {
        const index = appData.eventos.findIndex(e => e.id == id);
        if(index > -1) {
            appData.eventos[index] = { id: parseInt(id), fecha, hora, titulo, ubicacion };
        }
        document.getElementById('btn-guardar-evento').innerText = "Guardar Evento";
        document.getElementById('btn-cancelar-edicion').classList.add('hidden');
    } else {
        appData.eventos.push({ id: Date.now(), fecha, hora, titulo, ubicacion });
    }

    guardarDatosServidor();
    document.querySelectorAll('#evento-edit-id, #evento-fecha, #evento-hora, #evento-titulo, #evento-ubicacion').forEach(el => el.value = '');
    alert("Evento procesado correctamente");
}

function prepararEdicionEvento(id) {
    const evento = appData.eventos.find(e => e.id === id);
    if(evento) {
        document.getElementById('evento-edit-id').value = evento.id;
        document.getElementById('evento-fecha').value = evento.fecha;
        document.getElementById('evento-hora').value = evento.hora;
        document.getElementById('evento-titulo').value = evento.titulo;
        document.getElementById('evento-ubicacion').value = evento.ubicacion;
        
        document.getElementById('btn-guardar-evento').innerText = "Actualizar Evento";
        document.getElementById('btn-cancelar-edicion').classList.remove('hidden');
        cerrarModal('modal-calendario');
    }
}

function cancelarEdicionEvento() {
    document.querySelectorAll('#evento-edit-id, #evento-fecha, #evento-hora, #evento-titulo, #evento-ubicacion').forEach(el => el.value = '');
    document.getElementById('btn-guardar-evento').innerText = "Guardar Evento";
    document.getElementById('btn-cancelar-edicion').classList.add('hidden');
}

function eliminarEvento(id) {
    if(confirm("¿Seguro que deseas eliminar este evento?")) {
        appData.eventos = appData.eventos.filter(e => e.id !== id);
        guardarDatosServidor();
        document.getElementById('evento-detalles-box').classList.add('hidden');
    }
}

// ==========================================
// 5. GESTIÓN DE ANEXOS
// ==========================================
function agregarAnexo() {
    const nombre = document.getElementById('anexo-nombre').value;
    let url = document.getElementById('anexo-url').value;
    
    if(!nombre || !url) return alert("El nombre y la URL son obligatorios");
    if (!url.startsWith('http')) url = 'https://' + url;

    appData.anexos.push({ id: Date.now(), nombre, url, visible: true });
    guardarDatosServidor();
    document.getElementById('anexo-nombre').value = ''; 
    document.getElementById('anexo-url').value = '';
}

function alternarVisibilidad(id) {
    let indice = appData.anexos.findIndex(a => a.id === id);
    if(indice > -1) { 
        appData.anexos[indice].visible = !appData.anexos[indice].visible; 
        guardarDatosServidor(); 
    }
}

function eliminarAnexo(id) {
    if(confirm("¿Seguro que deseas eliminar este anexo?")) {
        appData.anexos = appData.anexos.filter(a => a.id !== id);
        guardarDatosServidor();
    }
}

function renderizarAnexos() {
    const listaAdmin = document.getElementById('lista-anexos-admin');
    const listaPublica = document.getElementById('lista-anexos-publico');
    
    if(listaAdmin) listaAdmin.innerHTML = ''; 
    if(listaPublica) listaPublica.innerHTML = '';
    
    if (appData.anexos.length === 0 && listaPublica) {
        listaPublica.innerHTML = '<li>No hay anexos disponibles en este momento.</li>';
    }

    appData.anexos.forEach(a => {
        // Render en Panel de Administrador
        if(listaAdmin) {
            listaAdmin.innerHTML += `
                <li>
                    <span>${a.nombre}</span> 
                    <div>
                        <button onclick="alternarVisibilidad(${a.id})" style="background:${a.visible?'green':'gray'}; color:white; padding:5px; width:auto; margin-right:5px;">
                            ${a.visible ? 'Visible' : 'Oculto'}
                        </button>
                        <button onclick="eliminarAnexo(${a.id})" class="btn-danger" style="padding:5px; width:auto;">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </div>
                </li>`;
        }
            
        // Render en Vista Pública
        if(a.visible && listaPublica) {
            listaPublica.innerHTML += `
                <li>
                    <a href="${a.url}" target="_blank" rel="noopener noreferrer">
                        <i class="fa-solid fa-file-pdf"></i> ${a.nombre}
                    </a>
                </li>`;
        }
    });
}