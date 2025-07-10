const clientId = "64713983477-nk4rmn95cgjsnab4gmp44dpjsdp1brk2.apps.googleusercontent.com";
const SPREADSHEET_ID = "1T8YifEIUU7a6ugf_Xn5_1edUUMoYfM9loDuOQU1u2-8"; // ID de tu Google Sheet
const SHEET_NAME_OBAMACARE = "Pólizas"; // Nombre de la hoja principal
const SHEET_NAME_CIGNA = "Cigna Complementario"; // Nombre de la hoja para Cigna
const SHEET_NAME_PAGOS = "Pagos"; // Nombre de la hoja para Pagos

const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets';

let tokenClient;
let accessToken = null;
let gapiInitialized = false;
let statusTimeout; 

// Elementos del DOM - Botones y Contenedores principales
const loginBtn = document.getElementById("loginBtn");
const dataForm = document.getElementById("dataForm");
const submitBtn = document.getElementById("submitBtn");
const fileInput = document.getElementById("fileInput");
const statusDiv = document.getElementById("status");

// Elementos para las pestañas
const tabButtons = document.querySelectorAll('.tabs-nav .tab-button');
const tabContents = document.querySelectorAll('.tab-content');

// Elementos del DOM - Campos de Obamacare
const cantidadDependientesInput = document.getElementById("cantidadDependientes");
const addDependentsBtn = document.getElementById("addDependentsBtn"); // Botón para abrir modal
const editDependentsBtn = document.getElementById("editDependentsBtn"); // Botón de edición de dependientes
const hasPoBoxCheckbox = document.getElementById('hasPoBox');
const poBoxAddressContainer = document.getElementById('poBoxAddressContainer');
const direccionCalleInput = document.getElementById('direccionCalle');
const casaApartamentoInput = document.getElementById('casaApartamento');
const condadoInput = document.getElementById('condado');
const ciudadInput = document.getElementById('ciudad');
const zipCodeInput = document.getElementById('zipCode');
const direccionPoBoxInput = document.getElementById('direccionPoBox');
const socialInput = document.getElementById('social');
const ingresosInput = document.getElementById('ingresos');
const creditoFiscalInput = document.getElementById('creditoFiscal');
const primaInput = document.getElementById('prima');

// Elementos del DOM - Modal de Dependientes
const dependentsModal = document.getElementById('dependentsModal');
const closeModalButton = dependentsModal.querySelector('.close-button');
const modalDependentsContainer = document.getElementById('modalDependentsContainer');
const saveDependentsBtn = document.getElementById('saveDependentsBtn');

// Elementos del DOM - Pestaña Pagos
const pagoBancoRadio = document.getElementById('pagoBanco');
const pagoTarjetaRadio = document.getElementById('pagoTarjeta');
const pagoBancoContainer = document.getElementById('pagoBancoContainer');
const pagoTarjetaContainer = document.getElementById('pagoTarjetaContainer');

// Campos de Tarjeta para eventos (no se guardan completos en Sheets)
const numTarjetaInput = document.getElementById('numTarjeta');
const fechaVencimientoInput = document.getElementById('fechaVencimiento');
const cvcInput = document.getElementById('cvc'); 
const titularTarjetaInput = document.getElementById('titularTarjeta'); // Asegurarse de tener la referencia


// Variable para almacenar dependientes temporalmente
let currentDependentsData = []; 
let currentTitularClientId = null; // Para vincular datos entre hojas

// Funciones de utilidad
/**
 * Muestra un mensaje de estado en la interfaz de usuario.
 * @param {string} msg - El mensaje a mostrar.
 * @param {'info'|'success'|'error'} type - El tipo de mensaje (para estilos CSS).
 * @param {number} duration - Duración en milisegundos para que el mensaje sea visible (0 para indefinido).
 */
function displayStatus(msg, type = 'info', duration = 5000) {
    clearTimeout(statusTimeout); 
    statusDiv.textContent = msg;
    statusDiv.className = `visible ${type}`; 
    
    if (duration > 0) {
        statusTimeout = setTimeout(() => {
            statusDiv.className = ''; 
            statusDiv.textContent = '';
        }, duration);
    }
}

/**
 * Formatea un valor numérico como moneda (ej: $123.45) al perder el foco.
 * Elimina caracteres no numéricos excepto el punto decimal.
 * @param {HTMLInputElement} inputElement - El elemento input a formatear.
 */
function formatCurrencyInput(inputElement) {
    let value = inputElement.value.replace(/[^0-9.]/g, ''); 
    const parts = value.split('.');
    if (parts.length > 2) {
        value = parts[0] + '.' + parts.slice(1).join('');
    }
    let floatValue = parseFloat(value);
    if (isNaN(floatValue) || value.trim() === '') {
        inputElement.value = ''; 
        return;
    }
    inputElement.value = `$${floatValue.toFixed(2)}`;
}

/**
 * Elimina el formato de moneda de un input para facilitar la edición.
 * @param {HTMLInputElement} inputElement - El elemento input.
 */
function unformatCurrencyInput(inputElement) {
    inputElement.value = inputElement.value.replace(/[^0-9.]/g, '');
}

/**
 * Convierte un valor de input formateado (ej. "$1,234.56") a un número limpio.
 * @param {string} formattedValue - El valor de la cadena formateada.
 * @returns {number} El valor numérico limpio.
 */
function parseFormattedMonetaryValue(formattedValue) {
    const cleanValue = formattedValue.replace(/[^0-9.]/g, '');
    return parseFloat(cleanValue);
}

// Lógica de Inicialización de GAPI
async function initGapiClient() {
    try {
        await gapi.client.init({
            discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4", "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"] 
        });
        await gapi.client.load('sheets', 'v4');
        await gapi.client.load('drive', 'v3'); 
        
        gapiInitialized = true;
        console.log("gapi.client inicializado y APIs de Sheets/Drive cargadas.");
    } catch (error) {
        console.error("Error al inicializar gapi.client para Sheets/Drive API:", error);
        displayStatus("Error crítico: No se pudo inicializar la API de Sheets/Drive. Revisa la consola.", 'error', 10000);
    }
}
gapi.load('client', initGapiClient);

window.onload = () => {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPES,
        callback: (response) => {
            if (response.error) {
                console.error("Error de autenticación:", response);
                displayStatus("Error al autenticar: " + response.error, 'error', 10000);
            } else {
                accessToken = response.access_token;
                displayStatus("Conectado con Google. Ya puedes enviar el formulario.", 'success');
                submitBtn.disabled = false; 
                loginBtn.disabled = true; 
            }
        },
    });

    // Ocultar contenedores de pago al cargar la página
    pagoBancoContainer.style.display = 'none';
    pagoTarjetaContainer.style.display = 'none';
};

// Event Listeners de la UI
loginBtn.addEventListener("click", () => {
    tokenClient.requestAccessToken();
});

// Lógica para mostrar/ocultar campo de PO BOX
hasPoBoxCheckbox.addEventListener('change', () => {
    if (hasPoBoxCheckbox.checked) {
        poBoxAddressContainer.style.display = 'block';
    } else {
        poBoxAddressContainer.style.display = 'none';
        direccionPoBoxInput.value = ''; 
    }
});

// Lógica para formato de SSN (en blur/focus)
socialInput.addEventListener('blur', function(e) {
    let value = e.target.value.replace(/\D/g, ''); 
    let formattedValue = '';
    if (value.length > 0) {
        formattedValue = value.substring(0, 3);
        if (value.length > 3) { formattedValue += '-' + value.substring(3, 5); }
        if (value.length > 5) { formattedValue += '-' + value.substring(5, 9); }
    }
    e.target.value = formattedValue;
});
socialInput.addEventListener('focus', function(e) {
    e.target.value = e.target.value.replace(/\D/g, ''); 
});

// Añadir listeners de 'blur' y 'focus' para formatear/desformatear campos monetarios
ingresosInput.addEventListener('blur', () => formatCurrencyInput(ingresosInput));
creditoFiscalInput.addEventListener('blur', () => formatCurrencyInput(creditoFiscalInput));
primaInput.addEventListener('blur', () => formatCurrencyInput(primaInput));
ingresosInput.addEventListener('focus', () => unformatCurrencyInput(ingresosInput));
creditoFiscalInput.addEventListener('focus', () => unformatCurrencyInput(creditoFiscalInput));
primaInput.addEventListener('focus', () => unformatCurrencyInput(primaInput));


// Lógica de Pestañas
tabButtons.forEach(button => {
    button.addEventListener('click', () => {
        const tabId = button.dataset.tab;

        // Ocultar todos los contenidos de las pestañas
        tabContents.forEach(content => content.style.display = 'none');
        // Remover clase 'active' de todos los botones
        tabButtons.forEach(btn => btn.classList.remove('active'));

        // Mostrar el contenido de la pestaña seleccionada
        document.getElementById(`tab-${tabId}`).style.display = 'block';
        // Añadir clase 'active' al botón seleccionado
        button.classList.add('active');
    });
});

// Lógica para mostrar/ocultar campos de pago (Banco/Tarjeta)
pagoBancoRadio.addEventListener('change', () => {
    pagoBancoContainer.style.display = pagoBancoRadio.checked ? 'grid' : 'none';
    // Limpiar campos de tarjeta si se selecciona banco
    numTarjetaInput.value = ''; fechaVencimientoInput.value = ''; cvcInput.value = ''; titularTarjetaInput.value = '';
    pagoTarjetaContainer.style.display = 'none'; 
});
pagoTarjetaRadio.addEventListener('change', () => {
    pagoTarjetaContainer.style.display = pagoTarjetaRadio.checked ? 'grid' : 'none';
    // Limpiar campos de banco si se selecciona tarjeta
    document.getElementById('numCuenta').value = ''; document.getElementById('numRuta').value = ''; 
    document.getElementById('nombreBanco').value = ''; document.getElementById('titularCuenta').value = ''; 
    document.getElementById('socialCuenta').value = '';
    pagoBancoContainer.style.display = 'none'; 
});


// Lógica del Modal de Dependientes
addDependentsBtn.addEventListener('click', () => openDependentsModal(false)); // Abrir para agregar/primera vez
editDependentsBtn.addEventListener('click', () => openDependentsModal(true)); // Abrir para editar

cantidadDependientesInput.addEventListener('change', () => {
    const num = parseInt(cantidadDependientesInput.value);
    
    // Validar y controlar visibilidad de botones
    if (isNaN(num) || num < 0) { 
        cantidadDependientesInput.value = 0; 
        displayStatus("Por favor, introduce una cantidad válida (número positivo) de dependientes.", 'error');
        currentDependentsData = [];
        addDependentsBtn.style.display = 'none';
        editDependentsBtn.style.display = 'none';
        return;
    }

    if (num > 0) {
        // Si ya hay datos y se ajusta el número, ajustar el array para precargar o añadir
        if (currentDependentsData.length !== num) {
            currentDependentsData = currentDependentsData.slice(0, num);
            for(let i = currentDependentsData.length; i < num; i++) {
                currentDependentsData.push({}); 
            }
        }
        // Mostrar el botón de editar si ya se guardaron dependientes antes o el de agregar si es nuevo
        if (currentDependentsData.some(dep => Object.keys(dep).length > 0)) { // Si ya hay algún dato en currentDependentsData
            editDependentsBtn.style.display = 'block';
            addDependentsBtn.style.display = 'none';
        } else {
            editDependentsBtn.style.display = 'none';
            addDependentsBtn.style.display = 'block';
        }
        
        // Abrir el modal automáticamente si es la primera vez que se pone >0 o si se acaba de ajustar el número y está vacío
        if ((num > 0 && currentDependentsData.length === 0) || (num > 0 && currentDependentsData.some(dep => Object.keys(dep).length === 0))) {
             openDependentsModal(false); // Abrir para agregar si no hay datos o hay vacíos
        }

    } else { // num === 0
        currentDependentsData = []; 
        addDependentsBtn.style.display = 'none';
        editDependentsBtn.style.display = 'none';
    }
});

closeModalButton.addEventListener('click', () => {
    dependentsModal.style.display = 'none';
});

saveDependentsBtn.addEventListener('click', () => {
    const num = parseInt(cantidadDependientesInput.value);
    for (let i = 0; i < num; i++) {
        const parentesco = document.getElementById(`modal_dep${i}_parentesco`).value;
        const nombre = document.getElementById(`modal_dep${i}_nombre`).value;
        const apellido = document.getElementById(`modal_dep${i}_apellido`).value;
        if (!parentesco || !nombre || !apellido) {
            displayStatus(`Faltan campos requeridos para el Dependiente #${i+1}.`, 'error');
            return; 
        }
    }

    saveDependentsFromModal();
    dependentsModal.style.display = 'none';
    editDependentsBtn.style.display = 'block'; 
    addDependentsBtn.style.display = 'none'; // Asegurarse que Agregar está oculto
});

// Cierra el modal si se hace clic fuera de su contenido
window.addEventListener('click', (event) => {
    if (event.target === dependentsModal) {
        dependentsModal.style.display = 'none';
    }
});

// Lógica de máscara para Número de Tarjeta y Fecha de Vencimiento
numTarjetaInput.addEventListener('input', function(e) {
    let value = e.target.value.replace(/\D/g, ''); 
    let formattedValue = '';
    for (let i = 0; i < value.length; i++) {
        if (i > 0 && i % 4 === 0) {
            formattedValue += ' ';
        }
        formattedValue += value[i];
    }
    e.target.value = formattedValue;
});

fechaVencimientoInput.addEventListener('input', function(e) {
    let value = e.target.value.replace(/\D/g, ''); 
    if (value.length > 2) {
        value = value.substring(0, 2) + '/' + value.substring(2, 4);
    }
    e.target.value = value;
});


/**
 * Abre el modal de dependientes.
 * @param {boolean} forEdit - True si se abre para editar datos existentes.
 */
function openDependentsModal(forEdit) {
    const num = parseInt(cantidadDependientesInput.value);
    if (isNaN(num) || num <= 0) {
        displayStatus("Introduce una cantidad válida y mayor a 0 para dependientes.", 'error');
        cantidadDependientesInput.value = 0; 
        addDependentsBtn.style.display = 'none'; // Ocultar botón si no hay dependientes
        editDependentsBtn.style.display = 'none';
        return;
    }

    modalDependentsContainer.innerHTML = ''; 

    // Ajustar `currentDependentsData` al número actual de dependientes si no coincide
    if (currentDependentsData.length !== num) {
        currentDependentsData = currentDependentsData.slice(0, num);
        for (let i = currentDependentsData.length; i < num; i++) {
            currentDependentsData.push({}); 
        }
    }
    
    // Generar los campos
    for (let i = 0; i < num; i++) {
        const dep = currentDependentsData[i] || {}; 
        const div = document.createElement("div");
        div.className = "dependent-card"; // Apply dependent-card class for styling
        div.innerHTML = `
            <h4>Dependiente #${i + 1}</h4>
            <div class="grid-item">
                <label for="modal_dep${i}_parentesco">Parentesco:</label>
                <input type="text" id="modal_dep${i}_parentesco" value="${dep.parentesco || ''}" placeholder="Ej: Hijo, Cónyuge" required>
            </div>
            <div class="grid-item">
                <label for="modal_dep${i}_nombre">Nombre:</label>
                <input type="text" id="modal_dep${i}_nombre" value="${dep.nombre || ''}" placeholder="Nombre del dependiente" required>
            </div>
            <div class="grid-item">
                <label for="modal_dep${i}_apellido">Apellido:</label>
                <input type="text" id="modal_dep${i}_apellido" value="${dep.apellido || ''}" placeholder="Apellido del dependiente" required>
            </div>
            <div class="grid-item">
                <label for="modal_dep${i}_fechaNacimiento">Fecha de nacimiento:</label>
                <input type="date" id="modal_dep${i}_fechaNacimiento" value="${dep.fechaNacimiento || ''}">
            </div>
            <div class="grid-item">
                <label for="modal_dep${i}_estadoMigratorio">Estado migratorio:</label>
                <select id="modal_dep${i}_estadoMigratorio">
                    <option value="">Selecciona...</option>
                    <option value="Ciudadano" ${dep.estadoMigratorio === 'Ciudadano' ? 'selected' : ''}>Ciudadano</option>
                    <option value="Residente" ${dep.estadoMigratorio === 'Residente' ? 'selected' : ''}>Residente</option>
                    <option value="Permiso de Trabajo" ${dep.estadoMigratorio === 'Permiso de Trabajo' ? 'selected' : ''}>Permiso de Trabajo</option>
                    <option value="Asilo Político" ${dep.estadoMigratorio === 'Asilo Político' ? 'selected' : ''}>Asilo Político</option>
                    <option value="I-94" ${dep.estadoMigratorio === 'I-94' ? 'selected' : ''}>I-94</option>
                    <option value="Otro" ${dep.estadoMigratorio === 'Otro' ? 'selected' : ''}>Otro</option>
                </select>
            </div>
            <div class="grid-item">
                <label for="modal_dep${i}_aplica">Aplica:</label>
                <select id="modal_dep${i}_aplica">
                    <option value="Sí" ${dep.aplica === 'Sí' ? 'selected' : ''}>Sí</option>
                    <option value="No" ${dep.aplica === 'No' ? 'selected' : ''}>No</option>
                </select>
            </div>
        `;
        modalDependentsContainer.appendChild(div);
    }
    dependentsModal.style.display = 'flex'; 
}

/**
 * Guarda los datos de los dependientes desde el modal a la variable temporal.
 */
function saveDependentsFromModal() {
    currentDependentsData = []; 
    const num = parseInt(cantidadDependientesInput.value);

    for (let i = 0; i < num; i++) {
        const parentesco = document.getElementById(`modal_dep${i}_parentesco`).value;
        const nombre = document.getElementById(`modal_dep${i}_nombre`).value;
        const apellido = document.getElementById(`modal_dep${i}_apellido`).value;
        const fechaNacimiento = document.getElementById(`modal_dep${i}_fechaNacimiento`).value;
        const estadoMigratorio = document.getElementById(`modal_dep${i}_estadoMigratorio`).value;
        const aplica = document.getElementById(`modal_dep${i}_aplica`).value;

        currentDependentsData.push({
            parentesco, nombre, apellido, fechaNacimiento, estadoMigratorio, aplica
        });
    }
    displayStatus("Dependientes guardados temporalmente.", 'info');
}


// Event listener para el envío del formulario
dataForm.addEventListener("submit", async (e) => {
    e.preventDefault(); 

    if (!accessToken) {
        displayStatus("Por favor, inicia sesión con Google primero.", 'error');
        return;
    }
    
    if (!gapiInitialized || !gapi.client.sheets || !gapi.client.drive) {
        displayStatus("Las APIs de Google no están completamente listas. Intenta recargar la página o espera un momento.", 'error', 10000);
        console.error("gapi.client no está completamente inicializado. No se puede proceder.");
        return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = "Enviando...";
    displayStatus("Procesando...", 'info', 0); 

    try {
        currentTitularClientId = `CLIENT-${Date.now()}-${Math.floor(Math.random() * 1000)}`; // Generar ID para este envío

        // --- Recolectar datos de la pestaña Obamacare (Principal) ---
        const titularData = {};
        const getInputValue = (id) => document.getElementById(id).value;
        
        titularData.nombre = getInputValue('nombre');
        titularData.apellidos = getInputValue('apellidos');
        titularData.sexo = getInputValue('sexo');
        titularData.correo = getInputValue('correo');
        titularData.telefono = getInputValue('telefono');
        titularData.fechaNacimiento = getInputValue('fechaNacimiento');
        titularData.estadoMigratorio = getInputValue('estadoMigratorio');
        titularData.aplica = getInputValue('aplica');
        titularData.cantidadDependientes = getInputValue('cantidadDependientes');
        
        // Desformatear valores monetarios (eliminar '$' y luego parsear a number)
        titularData.ingresos = parseFormattedMonetaryValue(getInputValue('ingresos'));
        titularData.social = getInputValue('social'); 
        titularData.compania = getInputValue('compania');
        titularData.plan = getInputValue('plan');
        titularData.creditoFiscal = parseFormattedMonetaryValue(getInputValue('creditoFiscal'));
        titularData.prima = parseFormattedMonetaryValue(getInputValue('prima'));
        titularData.link = getInputValue('link');
        titularData.observaciones = getInputValue('observaciones');
        titularData.Fecha = new Date().toLocaleDateString('es-ES'); 
        titularData.clientId = currentTitularClientId;

        // Procesar la dirección consolidada
        let fullAddress = '';
        if (hasPoBoxCheckbox.checked) {
            fullAddress = direccionPoBoxInput.value.trim();
        } else {
            const calle = direccionCalleInput.value.trim();
            const casaApto = casaApartamentoInput.value.trim();
            const condado = condadoInput.value.trim();
            const ciudad = ciudadInput.value.trim();
            const zip = zipCodeInput.value.trim();

            const addressParts = [];
            if (calle) addressParts.push(calle);
            if (casaApto) addressParts.push(casaApto);
            if (condado) addressParts.push(condado);
            if (ciudad) addressParts.push(ciudad);
            if (zip) addressParts.push(zip);
            
            fullAddress = addressParts.join(', ');
        }
        titularData.direccion = fullAddress; 

        // --- Procesar documentos ---
        let fileHyperlinks = []; 
        if (fileInput.files.length > 0) {
            displayStatus(`Subiendo ${fileInput.files.length} archivo(s) a Drive...`, 'info', 0);
            for (const file of fileInput.files) {
                const driveResponse = await uploadFileToDrive(file);
                fileHyperlinks.push(`HYPERLINK("${driveResponse.webViewLink}", "${file.name}")`); 
            }
            titularData.documentos = '=' + fileHyperlinks.join(' & CHAR(10) & ');
        } else {
            titularData.documentos = ''; 
        }

        // --- Enviar datos a Obamacare (Hoja principal) ---
        // Aquí se formatean los valores monetarios numéricos a `$X.XX` para la hoja de cálculo
        titularData.ingresos = titularData.ingresos ? `$${titularData.ingresos.toFixed(2)}` : '';
        titularData.creditoFiscal = titularData.creditoFiscal ? `$${titularData.creditoFiscal.toFixed(2)}` : '';
        titularData.prima = titularData.prima ? `$${titularData.prima.toFixed(2)}` : '';

        await appendObamacareDataToSheet(titularData, currentDependentsData);


        // --- Recolectar y enviar datos de Cigna Complementario ---
        const cignaData = { clientId: currentTitularClientId };
        cignaData.planTipo = getInputValue('cignaPlanTipo');
        cignaData.coberturaTipo = getInputValue('cignaCoberturaTipo');
        cignaData.beneficio = getInputValue('cignaBeneficio');
        cignaData.deducible = getInputValue('cignaDeducible');
        cignaData.mensualidad = getInputValue('cignaMensualidad');
        
        // Solo envía si hay datos significativos de Cigna (excluyendo el clientId vacío)
        if (Object.values(cignaData).some(val => val !== '' && val !== currentTitularClientId)) {
            await appendCignaDataToSheet(cignaData);
        }

        // --- Recolectar y enviar datos de Pagos ---
        const paymentData = { clientId: currentTitularClientId };
        paymentData.metodoPago = document.querySelector('input[name="metodoPago"]:checked')?.value || '';

        if (paymentData.metodoPago === 'Banco') {
            paymentData.numCuenta = getInputValue('numCuenta');
            paymentData.numRuta = getInputValue('numRuta');
            paymentData.nombreBanco = getInputValue('nombreBanco');
            paymentData.titularCuenta = getInputValue('titularCuenta');
            paymentData.socialCuenta = getInputValue('socialCuenta');
            // Asegurarse de que los campos de tarjeta no se envíen
            paymentData.numTarjeta = ''; // No se envía la tarjeta completa
            paymentData.fechaVencimiento = '';
            paymentData.titularTarjeta = '';
        } else if (paymentData.metodoPago === 'Tarjeta') {
            const fullCardNum = getInputValue('numTarjeta').replace(/\D/g, '');
            paymentData.numTarjeta = fullCardNum.slice(-4); // ¡Solo se guardan los últimos 4 dígitos!
            paymentData.fechaVencimiento = getInputValue('fechaVencimiento');
            // CVV/CVC NO SE RECOLECTA NI SE ENVÍA A NINGUNA PARTE POR SEGURIDAD
            paymentData.titularTarjeta = getInputValue('titularTarjeta');
            // Asegurarse de que los campos de banco no se envíen
            paymentData.numCuenta = '';
            paymentData.numRuta = '';
            paymentData.nombreBanco = '';
            paymentData.titularCuenta = '';
            paymentData.socialCuenta = '';
        }
        
        // Solo envía si se seleccionó un método de pago
        if (paymentData.metodoPago) {
            await appendPaymentDataToSheet(paymentData);
        }


        displayStatus("✅ ¡Datos guardados en Sheets y archivo(s) subido(s) a Drive!", 'success', 8000);
        
        // Resetear todos los campos y estados después del envío exitoso
        dataForm.reset(); 
        fileInput.value = ''; 
        cantidadDependientesInput.value = 0;
        currentDependentsData = []; // Limpiar datos de dependientes en memoria
        editDependentsBtn.style.display = 'none'; // Ocultar botón editar dependientes
        addDependentsBtn.style.display = 'none'; // Ocultar botón agregar dependientes
        hasPoBoxCheckbox.checked = false;
        poBoxAddressContainer.style.display = 'none';
        pagoBancoRadio.checked = false;
        pagoTarjetaRadio.checked = false;
        pagoBancoContainer.style.display = 'none';
        pagoTarjetaContainer.style.display = 'none';
        
        // Volver a la primera pestaña
        tabButtons[0].click(); 

    } catch (error) {
        console.error("Error completo en el proceso:", error);
        displayStatus("Ocurrió un error al procesar tu solicitud: " + error.message, 'error', 10000);
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = "Enviar Datos Completos";
    }
});

/**
 * Sube un archivo a Google Drive y le da permisos públicos de lectura.
 * @param {File} file - El archivo a subir.
 * @returns {Promise<{webViewLink: string, id: string}>} - Un objeto con el enlace de visualización del archivo y su ID.
 */
async function uploadFileToDrive(file) {
    const metadata = { name: file.name, mimeType: file.type };
    const form = new FormData();
    form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
    form.append("file", file);

    const uploadResponse = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}` },
        body: form
    });
    if (!uploadResponse.ok) {
        const errorData = await uploadResponse.json();
        throw new Error("Error al subir archivo a Drive (fetch): " + (errorData.error?.message || uploadResponse.statusText));
    }
    const uploadResult = await uploadResponse.json();
    const fileId = uploadResult.id;

    await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
        method: "POST",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ role: "reader", type: "anyone" })
    });

    const fileInfoResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?fields=webViewLink`, {
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!fileInfoResponse.ok) {
        const errorData = await fileInfoResponse.json();
        throw new Error("Error al obtener webViewLink: " + (errorData.error?.message || fileInfoResponse.statusText));
    }
    const fileInfo = await fileInfoResponse.json();
    return { webViewLink: fileInfo.webViewLink, id: fileId };
}


/**
 * Añade los datos del titular y dependientes a la hoja principal "Pólizas".
 * @param {object} titularData - Objeto con los datos del titular.
 * @param {Array<object>} dependents - Array de objetos con los datos de los dependientes.
 */
async function appendObamacareDataToSheet(titularData, dependents) {
    gapi.client.setToken({ access_token: accessToken });

    const allRows = [];

    // Fila del titular para Obamacare
    const titularRow = [
        'Titular', 
        titularData.nombre || '',
        titularData.apellidos || '',
        titularData.sexo || '', 
        titularData.correo || '', 
        titularData.direccion || '', 
        titularData.telefono || '',
        titularData.fechaNacimiento || '',
        titularData.estadoMigratorio || '',
        titularData.aplica || '', 
        titularData.cantidadDependientes || '',
        // Valores monetarios ya vienen formateados con $ del submit handler
        titularData.ingresos || '', 
        titularData.social || '', 
        titularData.compania || '', 
        titularData.plan || '',
        titularData.creditoFiscal || '', 
        titularData.prima || '', 
        titularData.link || '',
        titularData.observaciones || '', 
        titularData.Fecha || '', 
        titularData.documentos || '', 
        titularData.clientId || '' 
    ];
    allRows.push(titularRow);

    // Filas de dependientes para Obamacare
    dependents.forEach(dep => {
        const dependentRow = [
            dep.parentesco || '', 
            dep.nombre || '',
            dep.apellido || '',
            '', '', '', '', // Vacío para Sexo, Correo, Dirección, Teléfono
            dep.fechaNacimiento || '',
            dep.estadoMigratorio || '',
            dep.aplica || '',
            '', '', '', '', '', '', '', '', '', '', '', // Vacío para Cantidad Dependientes, Ingresos, SSN, Compañía, Plan, Crédito Fiscal, Prima, Link, Observaciones, Fecha, Documentos
            dep.clientId || '' 
        ];
        allRows.push(dependentRow);
    });

    const params = {
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME_OBAMACARE}!A:Z`, 
        valueInputOption: "USER_ENTERED", 
        insertDataOption: "INSERT_ROWS", 
    };
    const requestBody = { values: allRows };

    try {
        const response = await gapi.client.sheets.spreadsheets.values.append(params, requestBody);
        console.log("Datos Obamacare/Pólizas escritos:", response);
    } catch (error) {
        console.error("Error al escribir en hoja Obamacare:", error);
        throw new Error("Error al escribir datos Obamacare: " + (error.result?.error?.message || error.message));
    }
}

/**
 * Añade los datos de Cigna Complementario a su hoja dedicada.
 * @param {object} cignaData - Objeto con los datos de Cigna.
 */
async function appendCignaDataToSheet(cignaData) {
    gapi.client.setToken({ access_token: accessToken });

    const values = [[
        cignaData.clientId || '',
        cignaData.planTipo || '',
        cignaData.coberturaTipo || '',
        cignaData.beneficio || '',
        cignaData.deducible || '',
        cignaData.mensualidad || ''
    ]];

    const params = {
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME_CIGNA}!A:F`, 
        valueInputOption: "USER_ENTERED", 
        insertDataOption: "INSERT_ROWS", 
    };
    const requestBody = { values: values };

    try {
        const response = await gapi.client.sheets.spreadsheets.values.append(params, requestBody);
        console.log("Datos Cigna Complementario escritos:", response);
    } catch (error) {
        console.error("Error al escribir en hoja Cigna Complementario:", error);
        throw new Error("Error al escribir datos Cigna: " + (error.result?.error?.message || error.message));
    }
}

/**
 * Añade los datos de Pagos a su hoja dedicada.
 * @param {object} paymentData - Objeto con los datos de pago.
 */
async function appendPaymentDataToSheet(paymentData) {
    gapi.client.setToken({ access_token: accessToken });

    // Los valores de tarjeta y CVV ya se manejan para no guardar datos sensibles.
    const values = [[
        paymentData.clientId || '',
        paymentData.metodoPago || '',
        paymentData.numCuenta || paymentData.numTarjeta || '', // Consolidado: Número de Cuenta o ÚLTIMOS 4 DÍGITOS de Tarjeta
        paymentData.numRuta || paymentData.fechaVencimiento || '', // Consolidado: Número de Ruta o Fecha Vencimiento
        paymentData.nombreBanco || paymentData.titularTarjeta || '', // Consolidado: Nombre Banco o Titular Tarjeta
        paymentData.socialCuenta || '' // Solo para cuenta bancaria
    ]];

    const params = {
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME_PAGOS}!A:F`, 
        valueInputOption: "USER_ENTERED", 
        insertDataOption: "INSERT_ROWS", 
    };
    const requestBody = { values: values };

    try {
        const response = await gapi.client.sheets.spreadsheets.values.append(params, requestBody);
        console.log("Datos Pagos escritos:", response);
    } catch (error) {
        console.error("Error al escribir en hoja Pagos:", error);
        throw new Error("Error al escribir datos de pagos: " + (error.result?.error?.message || error.message));
    }
}