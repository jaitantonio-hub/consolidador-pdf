/* taskpane.js */

// Base de datos de correos para simulación fuera de Outlook (Mock Mode)
const mockEmailsDatabase = {
    "msg-01": {
        subject: "Re: Cotización Proyecto de Migración de Datos",
        from: "Carlos Mendoza <carlos.mendoza@empresa.com>",
        to: ["Jaime Tarazona <jaime.tarazona@outlook.com>"],
        cc: ["Ana María Gómez <ana.gomez@corporativo.com>"],
        dateTimeCreated: new Date(Date.now() - 3600000 * 2), // hace 2 horas
        attachments: [
            { name: "propuesta_comercial_v2.pdf", size: 1887436 },
            { name: "cronograma_fase1.xlsx", size: 348160 }
        ],
        body: `
            <p>Hola Jaime,</p>
            <p>Te adjunto la propuesta comercial actualizada para el sistema de extracción y consolidación de correos a PDF. Hemos tenido en cuenta los comentarios de la reunión técnica de ayer con tu equipo.</p>
            <p>Por favor, revisa el archivo adjunto con el cronograma y confírmame si las fechas se adaptan a tus plazos actuales.</p>
            <p>Quedo atento a tus comentarios para proceder con el cierre del radicado correspondiente.</p>
            <p>Saludos cordiales,<br><strong>Carlos Mendoza</strong><br>Director de Proyectos | TechSolutions</p>
        `
    },
    "msg-02": {
        subject: "RE: Aprobación de Presupuesto - Radicado #2026-99",
        from: "Ana María Gómez <ana.gomez@corporativo.com>",
        to: ["Jaime Tarazona <jaime.tarazona@outlook.com>", "Carlos Mendoza <carlos.mendoza@empresa.com>"],
        cc: [],
        dateTimeCreated: new Date(Date.now() - 3600000 * 1), // hace 1 hora
        attachments: [],
        body: `
            <p>Estimado Jaime,</p>
            <p>He revisado el informe final y el presupuesto para el Radicado #2026-99 ha sido <strong>APROBADO</strong> formalmente por la gerencia de finanzas.</p>
            <p>Podemos proceder con la fase de implementación a partir del próximo lunes. Carlos, por favor prepara el ambiente de pruebas para el Add-in.</p>
            <p>Atentamente,<br><strong>Ana María Gómez</strong><br>Directora Financiera | Corporativo</p>
        `
    },
    "msg-03": {
        subject: "Acuse de Recibo Oficial - Documentación Legal Radicado #2026-99",
        from: "Notificaciones Legales <legal@seguro.com>",
        to: ["Jaime Tarazona <jaime.tarazona@outlook.com>"],
        cc: [],
        dateTimeCreated: new Date(Date.now() - 600000), // hace 10 minutos
        attachments: [
            { name: "contrato_firmado_digitalmente.pdf", size: 4718592 }
        ],
        body: `
            <div style="font-family: Arial, sans-serif; border: 1px solid #dcdcdc; padding: 16px; background-color: #fafafa; border-radius: 4px;">
                <h2 style="color: #0078d4; margin-top: 0; font-size: 16px; border-bottom: 2px solid #0078d4; padding-bottom: 8px;">ACUSE DE RECIBO AUTOMÁTICO</h2>
                <p>Le informamos que la Notaría y el Departamento Legal han registrado con éxito la firma del contrato de licenciamiento.</p>
                <p><strong>Detalles del Registro:</strong></p>
                <table style="width: 100%; font-size: 13px; border-collapse: collapse; margin-top: 8px;">
                    <tr>
                        <td style="padding: 4px 0; font-weight: bold; width: 120px;">ID Transacción:</td>
                        <td style="padding: 4px 0;">TX-900881B</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; font-weight: bold;">Estado:</td>
                        <td style="padding: 4px 0; color: #107c41; font-weight: bold;">Procesado y Validado</td>
                    </tr>
                </table>
                <hr style="border: 0; border-top: 1px solid #eee; margin: 12px 0;">
                <p style="font-size: 11px; color: #666; margin-bottom: 0;">Este es un mensaje generado de forma automática. Por favor no responda a este remitente.</p>
            </div>
        `
    }
};

// Listado de correos seleccionados (en memoria local)
let currentSelectedItems = [];
let isSimulationMode = false;

// Inicialización de la Aplicación de forma robusta
let isInitialized = false;

function initApp(info) {
    if (isInitialized) return;
    isInitialized = true;
    
    console.log("Inicializando aplicación...");
    
    // Configurar elementos del DOM
    const btnConsolidate = document.getElementById("btn-consolidate");
    const btnBack = document.getElementById("btn-back");
    
    if (btnConsolidate) btnConsolidate.addEventListener("click", startConsolidation);
    if (btnBack) btnBack.addEventListener("click", resetApp);

    // Detección segura de si se ejecuta dentro de Outlook
    // Si viene de Outlook, la URL contiene parámetros como "host" o "_host_info"
    const hasOutlookParams = window.location.search.toLowerCase().includes("host") || 
                             window.location.search.toLowerCase().includes("_host_info");
    
    const isOutlook = (info && info.host === Office.HostType.Outlook) || hasOutlookParams;
    
    if (isOutlook) {
        console.log("Detectado entorno oficial de Outlook. Iniciando Add-in...");
        initializeOutlookAddIn();
    } else {
        console.log("Detectado navegador estándar fuera de Outlook. Iniciando modo de simulación...");
        initializeSimulationMode();
    }
}

// 1. Intentar inicializar de forma estándar con Office.onReady
try {
    Office.onReady((info) => {
        initApp(info);
    });
} catch (e) {
    console.warn("Error al registrar Office.onReady, se usará el fallback:", e);
}

// 2. Fallback de seguridad para modo simulación (Solo si NO está cargando dentro de Outlook)
const hasOutlookParams = window.location.search.toLowerCase().includes("host") || 
                         window.location.search.toLowerCase().includes("_host_info");

if (!hasOutlookParams) {
    setTimeout(() => {
        if (!isInitialized) {
            console.log("Office.onReady no respondió en el tiempo límite. Activando modo simulación...");
            initApp(null);
        }
    }, 2500); // 2.5 segundos de margen en navegadores estándar
} else {
    // Si está en Outlook, no activamos simulación por tiempo de espera.
    // Solo mostramos advertencia si el SDK de Microsoft falla tras 12 segundos.
    setTimeout(() => {
        if (!isInitialized) {
            showError("El SDK de Office (Office.js) está tardando en cargar debido a la conexión. Por favor, recarga el panel.");
        }
    }, 12000);
}

// ==========================================================================
// Configuración de Modo Outlook Oficial
// ==========================================================================
function initializeOutlookAddIn() {
    // Comprobar compatibilidad con selección múltiple (Mailbox 1.13)
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
        showError("Tu versión de Outlook o cliente no soporta la selección múltiple de correos (Se requiere Mailbox 1.13+).");
        return;
    }

    // Suscribir manejador de eventos de cambio de selección
    Office.context.mailbox.addHandlerAsync(
        Office.EventType.SelectedItemsChanged, 
        onSelectionChanged, 
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Error al registrar handler de selección:", result.error.message);
            } else {
                // Ejecutar inmediatamente para cargar la selección actual al abrir el panel
                onSelectionChanged();
            }
        }
    );
}

// Manejador del cambio de selección de correos
function onSelectionChanged() {
    Office.context.mailbox.getSelectedItemsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Error al obtener correos seleccionados:", result.error.message);
            showError("No se pudieron leer los correos seleccionados de la bandeja.");
            return;
        }

        currentSelectedItems = result.value || [];
        updateSelectionUI();
    });
}

// ==========================================================================
// Configuración de Modo Simulación (Para Pruebas)
// ==========================================================================
function initializeSimulationMode() {
    isSimulationMode = true;
    document.getElementById("mock-banner").classList.remove("hidden");
    document.getElementById("options-panel").classList.remove("hidden");

    // En modo simulación, precargamos los 3 correos ficticios
    currentSelectedItems = [
        { itemId: "msg-01", subject: mockEmailsDatabase["msg-01"].subject, sender: "Carlos Mendoza" },
        { itemId: "msg-02", subject: mockEmailsDatabase["msg-02"].subject, sender: "Ana María Gómez" },
        { itemId: "msg-03", subject: mockEmailsDatabase["msg-03"].subject, sender: "Notificaciones Legales" }
    ];

    updateSelectionUI();
}

// ==========================================================================
// Lógica de Interfaz de Usuario (UI)
// ==========================================================================
function updateSelectionUI() {
    const listContainer = document.getElementById("email-list");
    const countBadge = document.getElementById("selected-count");
    const btnConsolidate = document.getElementById("btn-consolidate");
    const warningBanner = document.getElementById("limit-warning");
    const optionsPanel = document.getElementById("options-panel");

    // Limpiar contenedor
    listContainer.innerHTML = "";
    
    const count = currentSelectedItems.length;
    countBadge.innerText = count;

    if (count === 0) {
        // Mostrar Estado Vacío
        listContainer.innerHTML = `
            <div class="empty-state">
                <svg class="empty-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M3 8L10.8906 13.2604C11.5624 13.7083 12.4376 13.7083 13.1094 13.2604L21 8M5 19H19C20.1046 19 21 18.1046 21 17V7C21 5.89543 20.1046 5 19 5H5C3.89543 5 3 5.89543 3 7V17C3 18.1046 3.89543 19 5 19Z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
                <p>Ningún correo seleccionado</p>
                <small>Marca las casillas de verificación de hasta 100 correos en la bandeja de entrada para comenzar.</small>
            </div>
        `;
        btnConsolidate.disabled = true;
        warningBanner.classList.add("hidden");
        optionsPanel.classList.add("hidden");
    } else {
        // Mostrar listado de correos seleccionados
        currentSelectedItems.forEach((item) => {
            const card = document.createElement("div");
            card.className = "email-item-card";
            
            // Asunto
            const subjectDiv = document.createElement("div");
            subjectDiv.className = "email-item-subject";
            subjectDiv.innerText = item.subject || "Sin asunto";
            card.appendChild(subjectDiv);

            // Remitente (en simulado se llama sender, en real varía, por seguridad mostramos lo que venga)
            const senderDiv = document.createElement("div");
            senderDiv.className = "email-item-sender";
            senderDiv.innerText = item.sender || "Correo seleccionado";
            card.appendChild(senderDiv);

            listContainer.appendChild(card);
        });

        optionsPanel.classList.remove("hidden");

        // Reglas de habilitación
        if (count > 100) {
            btnConsolidate.disabled = true;
            warningBanner.classList.remove("hidden");
        } else {
            btnConsolidate.disabled = false;
            warningBanner.classList.add("hidden");
        }
    }
}

// ==========================================================================
// Extracción y Consolidación de Datos
// ==========================================================================
async function startConsolidation() {
    const selectionView = document.getElementById("selection-view");
    const loadingView = document.getElementById("loading-view");
    
    // Cambiar a vista de carga
    selectionView.classList.add("hidden");
    loadingView.classList.remove("hidden");

    const total = currentSelectedItems.length;
    const loadedEmails = [];

    updateProgressBar(0, total, "Iniciando descarga de datos...");

    try {
        for (let i = 0; i < total; i++) {
            const selectedItem = currentSelectedItems[i];
            updateProgressBar(i, total, `Procesando correo ${i + 1} de ${total}: "${selectedItem.subject}"`);

            let emailDetails;
            if (isSimulationMode) {
                // Cargar datos ficticios
                emailDetails = await getMockEmailDetails(selectedItem.itemId);
            } else {
                // Cargar datos reales usando loadItemByIdAsync de Office.js
                emailDetails = await loadRealEmailDetails(selectedItem.itemId);
            }

            loadedEmails.push(emailDetails);
        }

        updateProgressBar(total, total, "Generando documento unificado...");
        
        // Generar el PDF
        await generateUnifiedPDF(loadedEmails);

    } catch (error) {
        console.error("Fallo durante la consolidación:", error);
        showError(`Ocurrió un error al consolidar los correos: ${error.message || error}`);
    }
}

// Lógica de carga real de correos de Outlook
function loadRealEmailDetails(itemId) {
    return new Promise((resolve, reject) => {
        // Mailbox 1.15 requerido para loadItemByIdAsync
        if (!Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
            reject("Tu cliente de Outlook no soporta la carga de detalles de correos por ID (Se requiere Mailbox 1.15+).");
            return;
        }

        Office.context.mailbox.loadItemByIdAsync(itemId, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                reject(new Error(result.error.message));
                return;
            }

            const loadedItem = result.value;

            // Recopilar propiedades básicas del mensaje
            const subject = loadedItem.subject || "Sin asunto";
            const from = loadedItem.from ? `${loadedItem.from.displayName} <${loadedItem.from.emailAddress}>` : "Remitente desconocido";
            
            const toRecipients = [];
            if (loadedItem.to) {
                loadedItem.to.forEach(r => toRecipients.push(`${r.displayName} <${r.emailAddress}>`));
            }

            const ccRecipients = [];
            if (loadedItem.cc) {
                loadedItem.cc.forEach(r => ccRecipients.push(`${r.displayName} <${r.emailAddress}>`));
            }

            const dateStr = loadedItem.dateTimeCreated ? new Date(loadedItem.dateTimeCreated).toLocaleString() : "Sin fecha";

            const attachments = [];
            if (loadedItem.attachments && loadedItem.attachments.length > 0) {
                loadedItem.attachments.forEach(att => {
                    attachments.push({
                        name: att.name,
                        size: att.size
                    });
                });
            }

            // Cargar el cuerpo HTML del correo de forma asíncrona
            loadedItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
                let htmlBody = "";
                if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    htmlBody = bodyResult.value;
                } else {
                    htmlBody = `<p style="color:red;">No se pudo extraer el cuerpo del correo: ${bodyResult.error.message}</p>`;
                }

                // Descargar el item cargado en memoria para liberar recursos
                loadedItem.unloadAsync((unloadResult) => {
                    if (unloadResult.status === Office.AsyncResultStatus.Failed) {
                        console.warn("No se pudo descargar el item de memoria:", unloadResult.error.message);
                    }
                    
                    // Resolver con los datos consolidados del correo
                    resolve({
                        subject,
                        from,
                        to: toRecipients,
                        cc: ccRecipients,
                        date: dateStr,
                        attachments,
                        body: htmlBody
                    });
                });
            });
        });
    });
}

// Lógica de carga de correos simulados
function getMockEmailDetails(itemId) {
    return new Promise((resolve) => {
        setTimeout(() => {
            const data = mockEmailsDatabase[itemId];
            resolve({
                subject: data.subject,
                from: data.from,
                to: data.to,
                cc: data.cc,
                date: data.dateTimeCreated.toLocaleString(),
                attachments: data.attachments,
                body: data.body
            });
        }, 600); // Pequeña demora para simular la latencia de red
    });
}

// Actualizar Barra de Progreso
function updateProgressBar(current, total, text) {
    const fill = document.getElementById("progress-bar-fill");
    const progressText = document.getElementById("progress-text");
    const loaderTitle = document.getElementById("loader-title");
    
    loaderTitle.innerText = text;
    
    const percentage = total > 0 ? (current / total) * 100 : 0;
    fill.style.width = `${percentage}%`;
    progressText.innerText = `${current} / ${total} correos procesados`;
}

// ==========================================================================
// Renderizado y Exportación de PDF (html2pdf.js)
// ==========================================================================
function generateUnifiedPDF(emails) {
    return new Promise((resolve, reject) => {
        const renderArea = document.getElementById("pdf-render-area");
        
        // Leer opciones del panel
        const usePageBreak = document.getElementById("opt-page-break").checked;
        const showAttachments = document.getElementById("opt-show-attachments").checked;

        // Construir la estructura HTML del documento PDF (unificado por bloques independientes de primer nivel)
        let documentHtml = "";

        emails.forEach((email, index) => {
            const isFirst = index === 0;
            
            // Inyectar el helper oficial html2pdf__page-break antes de cada correo (excepto el primero)
            if (!isFirst && usePageBreak) {
                documentHtml += `<div class="html2pdf__page-break"></div>`;
            }

            const blockClass = "pdf-email-block";

            documentHtml += `
                <div class="pdf-document">
                    <div class="${blockClass}" style="margin-bottom: 30px;">
                        <!-- Cabecera Oficial Outlook -->
                        <div class="pdf-outlook-brand">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle; margin-right: 8px;">
                                <path d="M21.5 5.5H10.5C9.4 5.5 8.5 6.4 8.5 7.5V16.5C8.5 17.6 9.4 18.5 10.5 18.5H21.5C22.6 18.5 23.5 17.6 23.5 16.5V7.5C23.5 6.4 22.6 5.5 21.5 5.5Z" fill="#0078D4"/>
                                <path d="M21.5 7.5L16 11.5L10.5 7.5V9.5L16 13.5L21.5 9.5V7.5Z" fill="white"/>
                                <path d="M10.5 3.5H3.5C2.4 3.5 1.5 4.4 1.5 5.5V18.5C1.5 19.6 2.4 20.5 3.5 20.5H10.5C11.6 20.5 12.5 19.6 12.5 18.5V5.5C12.5 4.4 11.6 3.5 10.5 3.5Z" fill="#106EBE"/>
                                <path d="M4.5 7.5H7.5V9.5H5.5V11.5H7.5V13.5H5.5V14.5H7.5V16.5H4.5V7.5Z" fill="white"/>
                            </svg>
                            <span class="pdf-outlook-brand-text">Outlook</span>
                        </div>
                        
                        <div class="pdf-email-top-divider"></div>
                        
                        <div class="pdf-email-header">
                            <div class="pdf-email-subject">${escapeHtml(email.subject)}</div>
                            
                            <div class="pdf-email-subject-divider"></div>
                            
                            <div class="pdf-email-meta-row">
                                <span class="pdf-email-meta-label">Desde</span>
                                <span>${escapeHtml(email.from)}</span>
                            </div>
                            <div class="pdf-email-meta-row">
                                <span class="pdf-email-meta-label">Fecha</span>
                                <span>${escapeHtml(email.date)}</span>
                            </div>
                            <div class="pdf-email-meta-row">
                                <span class="pdf-email-meta-label">Para</span>
                                <span>${escapeHtml(email.to.join("; "))}</span>
                            </div>
            `;

            // Fila CC opcional
            if (email.cc && email.cc.length > 0) {
                documentHtml += `
                            <div class="pdf-email-meta-row">
                                <span class="pdf-email-meta-label">CC</span>
                                <span>${escapeHtml(email.cc.join("; "))}</span>
                            </div>
                `;
            }

            // Archivos adjuntos opcionales impresos como texto (Estilo Outlook)
            if (showAttachments && email.attachments && email.attachments.length > 0) {
                const count = email.attachments.length;
                const totalSize = email.attachments.reduce((acc, att) => acc + (att.size || 0), 0);
                const sizeStr = totalSize > 0 ? formatBytes(totalSize) : "";
                
                const labelAdjunto = count === 1 ? "archivo adjunto" : "archivos adjuntos";
                const sizeParen = sizeStr ? ` (${sizeStr})` : "";
                
                documentHtml += `
                            <div class="pdf-attachments-block">
                                <div class="pdf-attachments-summary">
                                    📎 ${count} ${labelAdjunto}${sizeParen}
                                </div>
                                <div class="pdf-attachments-filenames">
                `;
                
                email.attachments.forEach(att => {
                    documentHtml += `${escapeHtml(att.name)};<br>`;
                });
                
                documentHtml += `
                                </div>
                            </div>
                `;
            }

            documentHtml += `
                        </div>
                        <!-- Cuerpo del Mensaje -->
                        <div class="pdf-email-body">
                            ${email.body}
                        </div>
                    </div>
                </div>
            `;
        });

        // Generar un nombre de archivo por defecto
        let defaultFilename = "correos_consolidados.pdf";
        if (emails.length === 1) {
            // Si es un solo correo, usar su asunto limpio
            const cleanSubject = emails[0].subject
                .replace(/[^a-zA-Z0-9\s-_]/g, "")
                .trim()
                .substring(0, 50);
            defaultFilename = `${cleanSubject || "correo"}.pdf`;
        } else if (emails.length > 1) {
            defaultFilename = `Lote_de_${emails.length}_correos.pdf`;
        }

        // Opciones de configuración para html2pdf
        const opt = {
            margin:       25.4, // Margen Norma APA (7ª Edición): 2.54 cm (1 pulgada) en todos los bordes
            filename:     defaultFilename,
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { 
                scale: 2,           // Mayor calidad visual
                useCORS: true,      // Permitir imágenes externas si tienen CORS
                logging: false,
                width: 800,         // Forzar ancho de captura de 800px para evitar recortes a la derecha
                windowWidth: 800    // Forzar ancho virtual de 800px
            },
            jsPDF:        { unit: 'mm', format: 'letter', orientation: 'portrait' },
            pagebreak:    { 
                mode: ['css', 'legacy']
            }
        };

        // Ejecutar html2pdf pasando la cadena de texto HTML directamente
        html2pdf()
            .set(opt)
            .from(documentHtml)
            .save()
            .then(() => {
                showSuccessScreen(defaultFilename, emails.length);
                resolve();
            })
            .catch((err) => {
                reject(err);
            });
    });
}

// ==========================================================================
// Ayudantes de Formateo y Control de Errores
// ==========================================================================
function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function escapeHtml(text) {
    if (!text) return "";
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function showSuccessScreen(filename, count) {
    document.getElementById("loading-view").classList.add("hidden");
    const successView = document.getElementById("success-view");
    successView.classList.remove("hidden");
    
    document.getElementById("pdf-filename").innerText = filename;
    document.getElementById("pdf-count").innerText = count;
}

function showError(message) {
    // Restaurar vistas y mostrar mensaje de error en alert
    document.getElementById("loading-view").classList.add("hidden");
    document.getElementById("selection-view").classList.remove("hidden");
    
    // Crear dinámicamente un banner de error temporal
    const body = document.querySelector(".app-body");
    const existingError = document.getElementById("temp-error");
    if (existingError) existingError.remove();

    const errBanner = document.createElement("div");
    errBanner.id = "temp-error";
    errBanner.className = "alert alert-danger";
    errBanner.innerHTML = `<strong>Error:</strong> ${escapeHtml(message)}`;
    
    // Insertar al principio del body
    body.insertBefore(errBanner, body.firstChild);
    
    // Desaparecer después de 10 segundos
    setTimeout(() => {
        if (errBanner.parentNode) errBanner.remove();
    }, 10000);
}

function resetApp() {
    document.getElementById("success-view").classList.add("hidden");
    document.getElementById("selection-view").classList.remove("hidden");
    
    // Borrar banners de error
    const existingError = document.getElementById("temp-error");
    if (existingError) existingError.remove();
    
    updateSelectionUI();
}
