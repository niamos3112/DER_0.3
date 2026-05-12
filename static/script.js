// ==========================================
// CONFIGURACIÓN GLOBAL
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

// Escuchar cambios en el selector para mostrar el formulario correcto
document.addEventListener('DOMContentLoaded', () => {
    cargarDatosServidor();
    
    const dirTipo = document.getElementById('dir-tipo');
    if(dirTipo) {
        dirTipo.addEventListener('change', function() {
            if(this.value === 'incubadora') {
                document.getElementById('form-incubadora').classList.remove('hidden');
                document.getElementById('form-formulador').classList.add('hidden');
            } else {
                document.getElementById('form-incubadora').classList.add('hidden');
                document.getElementById('form-formulador').classList.remove('hidden');
            }
        });
    }
});

// ==========================================
// 1. COMUNICACIÓN CON EL SERVIDOR
// ==========================================
async function cargarDatosServidor() {
    try {
        const respuesta = await fetch(API_URL);
        if (respuesta.ok) {
            appData = await respuesta.json();
            const banner = document.getElementById('loading-banner');
            if(banner) {
                banner.style.opacity = '0';
                setTimeout(() => banner.style.display = 'none', 500);
            }
            renderizarTodo();
        }
    } catch (error) {
        console.error("Error conectando al servidor", error);
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
    const btnPostula = document.getElementById('btn-postula-publico');
    const inputLink = document.getElementById('admin-link-postula');
    if (btnPostula) btnPostula.href = appData.linkPostulacion;
    if (inputLink) inputLink.value = appData.linkPostulacion;
    
    renderizarEntidadesAdmin();
    renderizarCalendario();
    renderizarAnexos();
}

// ==========================================
// 2. CONTROL DE ACCESO
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
        renderizarCalendario();
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
    alert("Enlace actualizado");
}

// ==========================================
// 3. GESTIÓN DE DIRECTORIOS (DOBLE MODELO)
// ==========================================
function agregarDirectorioManual() {
    const tipo = document.getElementById('dir-tipo').value;
    const categoria = document.getElementById('dir-categoria').value;
    
    // Asignamos un ID único a cada registro para poder borrarlo sin conflictos
    let nuevoRegistro = { id: Date.now(), Tipo: tipo, Categoria: categoria };

    if(tipo === 'incubadora') {
        nuevoRegistro.Organizacion = document.getElementById('inc-org').value.trim();
        nuevoRegistro.RUC = document.getElementById('inc-ruc').value.trim();
        nuevoRegistro.Tipo_Entidad = document.getElementById('inc-tipo').value.trim();
        nuevoRegistro.Celular = document.getElementById('inc-cel').value.trim();
        nuevoRegistro.Correo = document.getElementById('inc-correo').value.trim();
        nuevoRegistro.Contacto_Nombres = document.getElementById('inc-nombres').value.trim();
        nuevoRegistro.Contacto_Cargo = document.getElementById('inc-cargo').value.trim();
        nuevoRegistro.Servicios = document.getElementById('inc-servicios').value.trim();
        nuevoRegistro.Sectores_Interes = document.getElementById('inc-sectores').value.trim();
        nuevoRegistro.Link_Portafolio = document.getElementById('inc-portafolio').value.trim();

        if(!nuevoRegistro.Organizacion) return alert("El nombre de la organización es obligatorio");
    } else {
        nuevoRegistro.Nombres = document.getElementById('form-nombres').value.trim();
        nuevoRegistro.Apellidos = document.getElementById('form-apellidos').value.trim();
        nuevoRegistro.Celular = document.getElementById('form-cel').value.trim();
        nuevoRegistro.Correo = document.getElementById('form-correo').value.trim();
        nuevoRegistro.Sectores_Priorizados = document.getElementById('form-sectores').value.trim();
        nuevoRegistro.Fondos_Experiencia = document.getElementById('form-fondos').value.trim();
        nuevoRegistro.Interes_Principal = document.getElementById('form-interes').value.trim();
        nuevoRegistro.Disponibilidad = document.getElementById('form-disp').value.trim();
        nuevoRegistro.Link_CV = document.getElementById('form-cv').value.trim();

        if(!nuevoRegistro.Nombres || !nuevoRegistro.Apellidos) return alert("Nombres y Apellidos son obligatorios");
    }

    appData.directorio.push(nuevoRegistro);
    guardarDatosServidor();
    
    // Limpiar inputs
    document.querySelectorAll('.admin-sub-section input').forEach(el => el.value = '');
    alert(`Registro añadido a ${categoria}`);
}

function descargarPlantilla() {
    const tipo = document.getElementById('dir-tipo').value;
    let encabezados = [];
    
    if(tipo === 'incubadora') {
        encabezados = [["Organizacion", "Celular", "Correo", "RUC", "Tipo_Entidad", "Contacto_Nombres", "Contacto_Cargo", "Servicios", "Sectores_Interes", "Link_Portafolio"]];
    } else {
        encabezados = [["Nombres", "Apellidos", "Celular", "Correo", "Sectores_Priorizados", "Fondos_Experiencia", "Interes_Principal", "Disponibilidad", "Link_CV"]];
    }

    const ws = XLSX.utils.aoa_to_sheet(encabezados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Datos");
    XLSX.writeFile(wb, `Plantilla_${tipo}.xlsx`);
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
        
        const datosMapeados = datosJSON.map(fila => {
            let obj = { id: Date.now() + Math.random(), Tipo: tipoSeleccionado, Categoria: categoriaSeleccionada };
            
            if(tipoSeleccionado === 'incubadora') {
                obj.Organizacion = (fila.Organizacion || '').toString().trim();
                obj.Celular = (fila.Celular || '').toString().trim();
                obj.Correo = (fila.Correo || '').toString().trim();
                obj.RUC = (fila.RUC || '').toString().trim();
                obj.Tipo_Entidad = (fila.Tipo_Entidad || '').toString().trim();
                obj.Contacto_Nombres = (fila.Contacto_Nombres || '').toString().trim();
                obj.Contacto_Cargo = (fila.Contacto_Cargo || '').toString().trim();
                obj.Servicios = (fila.Servicios || '').toString().trim();
                obj.Sectores_Interes = (fila.Sectores_Interes || '').toString().trim();
                obj.Link_Portafolio = (fila.Link_Portafolio || '').toString().trim();
            } else {
                obj.Nombres = (fila.Nombres || '').toString().trim();
                obj.Apellidos = (fila.Apellidos || '').toString().trim();
                obj.Celular = (fila.Celular || '').toString().trim();
                obj.Correo = (fila.Correo || '').toString().trim();
                obj.Sectores_Priorizados = (fila.Sectores_Priorizados || '').toString().trim();
                obj.Fondos_Experiencia = (fila.Fondos_Experiencia || '').toString().trim();
                obj.Interes_Principal = (fila.Interes_Principal || '').toString().trim();
                obj.Disponibilidad = (fila.Disponibilidad || '').toString().trim();
                obj.Link_CV = (fila.Link_CV || '').toString().trim();
            }
            return obj;
        });
        
        appData.directorio = [...appData.directorio, ...datosMapeados];
        guardarDatosServidor();
        document.getElementById('excel-input').value = ''; 
        alert(`Excel cargado correctamente como ${tipoSeleccionado}`);
    };
    reader.readAsArrayBuffer(archivo);
}

function eliminarRegistroDirectorio(id) {
    if(confirm(`¿Seguro que deseas eliminar este registro?`)) {
        appData.directorio = appData.directorio.filter(d => d.id !== id);
        guardarDatosServidor();
    }
}

function renderizarEntidadesAdmin() {
    const lista = document.getElementById('lista-admin-entidades');
    if(!lista) return;

    lista.innerHTML = ''; 
    
    appData.directorio.forEach(item => {
        const li = document.createElement('li');
        let tituloPrincipal = item.Tipo === 'incubadora' ? item.Organizacion : `${item.Nombres} ${item.Apellidos}`;
        
        li.innerHTML = `<span><strong>${tituloPrincipal}</strong> <br><small style="color:#555;">(${item.Tipo} - ${item.Categoria})</small></span>`;
        
        const btnEliminar = document.createElement('button');
        btnEliminar.className = 'btn-danger';
        btnEliminar.innerHTML = '<i class="fa-solid fa-trash"></i>';
        btnEliminar.onclick = () => eliminarRegistroDirectorio(item.id);
        
        li.appendChild(btnEliminar);
        lista.appendChild(li);
    });
}

function abrirDirectorio(tipo, categoria) {
    document.getElementById('titulo-directorio').innerText = tipo === 'incubadora' ? 'Incubadoras y Otras Entidades' : 'Formuladores';
    
    const data = appData.directorio.filter(d => d.Tipo === tipo && d.Categoria === categoria);
    const contenedor = document.getElementById('contenedor-entidades');
    contenedor.innerHTML = '';

    if(data.length === 0) {
        contenedor.innerHTML = '<p>No hay registros en esta sección para esta página.</p>';
        abrirModal('modal-directorio');
        return;
    }

    if(tipo === 'incubadora') {
        // Vista para Incubadoras (Fichas por organización)
        let htmlFichas = data.map(i => {
            let linkPortafolio = i.Link_Portafolio ? `<a href="${i.Link_Portafolio.startsWith('http') ? i.Link_Portafolio : 'https://'+i.Link_Portafolio}" target="_blank">Ver Portafolio</a>` : '-';
            return `
            <div class="entidad-group">
                <div class="entidad-header"><i class="fa-solid fa-building"></i> ${i.Organizacion || '-'}</div>
                <div class="table-responsive">
                    <table>
                        <tbody>
                            <tr>
                                <th>RUC</th><td data-label="RUC">${i.RUC || '-'}</td>
                                <th>Tipo Entidad</th><td data-label="Tipo Entidad">${i.Tipo_Entidad || '-'}</td>
                            </tr>
                            <tr>
                                <th>Contacto</th><td data-label="Contacto">${i.Contacto_Nombres || '-'}</td>
                                <th>Cargo</th><td data-label="Cargo">${i.Contacto_Cargo || '-'}</td>
                            </tr>
                            <tr>
                                <th>Celular</th><td data-label="Celular">${i.Celular || '-'}</td>
                                <th>Correo</th><td data-label="Correo">${i.Correo || '-'}</td>
                            </tr>
                            <tr>
                                <th>Servicios</th><td colspan="3" data-label="Servicios">${i.Servicios || '-'}</td>
                            </tr>
                            <tr>
                                <th>Sectores Interés</th><td colspan="3" data-label="Sectores">${i.Sectores_Interes || '-'}</td>
                            </tr>
                            <tr>
                                <th>Portafolio</th><td colspan="3" data-label="Portafolio">${linkPortafolio}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>`;
        }).join('');
        contenedor.innerHTML = htmlFichas;

    } else {
        // Vista para Formuladores (Tabla Plana)
        let filas = data.map(i => {
            let linkCV = i.Link_CV ? `<a href="${i.Link_CV.startsWith('http') ? i.Link_CV : 'https://'+i.Link_CV}" target="_blank">Ver CV</a>` : '-';
            return `
            <tr>
                <td data-label="Nombre">${i.Nombres} ${i.Apellidos}</td>
                <td data-label="Contacto">${i.Celular}<br>${i.Correo}</td>
                <td data-label="Sectores">${i.Sectores_Priorizados || '-'}</td>
                <td data-label="Fondos">${i.Fondos_Experiencia || '-'}</td>
                <td data-label="Interés">${i.Interes_Principal || '-'}</td>
                <td data-label="Disponibilidad">${i.Disponibilidad || '-'}</td>
                <td data-label="CV">${linkCV}</td>
            </tr>`;
        }).join('');
        
        contenedor.innerHTML = `
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Nombre Completo</th>
                            <th>Contacto</th>
                            <th>Sectores Priorizados</th>
                            <th>Fondos de Experiencia</th>
                            <th>Interés Principal</th>
                            <th>Disponibilidad</th>
                            <th>CV</th>
                        </tr>
                    </thead>
                    <tbody>${filas}</tbody>
                </table>
            </div>`;
    }
    
    abrirModal('modal-directorio');
}

// ==========================================
// 4. CALENDARIO Y EVENTOS
// ==========================================
function cambiarMes(offset) {
    fechaActual.setMonth(fechaActual.getMonth() + offset);
    renderizarCalendario();
}

function renderizarCalendario() {
    const grid = document.getElementById('calendar-body');
    const displayMes = document.getElementById('mes-anio-display');
    const boxDetalles = document.getElementById('evento-detalles-box');
    
    if(!grid || !displayMes) return; 

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
            ubicacionHTML = `<a href="${urlSegura}" target="_blank" rel="noopener noreferrer" class="btn-meet-link"><i class="fa-solid fa-video"></i> Ir a la Reunión</a>`;
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
        if(index > -1) appData.eventos[index] = { id: parseInt(id), fecha, hora, titulo, ubicacion };
        document.getElementById('btn-guardar-evento').innerText = "Guardar Evento";
        document.getElementById('btn-cancelar-edicion').classList.add('hidden');
    } else {
        appData.eventos.push({ id: Date.now(), fecha, hora, titulo, ubicacion });
    }

    guardarDatosServidor();
    document.querySelectorAll('#evento-edit-id, #evento-fecha, #evento-hora, #evento-titulo, #evento-ubicacion').forEach(el => el.value = '');
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