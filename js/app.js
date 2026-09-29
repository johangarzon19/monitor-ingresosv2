    // ============================================================
    // MONITOR DE INGRESOS
    // ============================================================

    let datosProgramacion = [];
    let datosIngresos = [];
    let resultados = [];


    // ============================================================
    // COMPONENTE: LISTA DESPLEGABLE CON SELECCIÓN MÚLTIPLE
    // ============================================================

    // Estilos base del filtro múltiple (por si el CSS está en caché)
    (function () {
        if (document.getElementById("ms-estilos-base")) return;
        const st = document.createElement("style");
        st.id = "ms-estilos-base";
        st.textContent =
            ".multiselect{position:relative;width:100%}" +
            ".ms-panel{position:absolute;top:calc(100% + 4px);left:0;z-index:50;width:100%;min-width:240px;padding:10px;background:#fff;border:1px solid #d1d5db;border-radius:8px;box-shadow:0 10px 25px rgba(0,0,0,.15)}" +
            ".ms-panel[hidden]{display:none!important}" +
            ".ms-lista{max-height:260px;overflow-y:auto}" +
            ".ms-opcion{display:flex;align-items:center;gap:8px;padding:6px 4px;cursor:pointer}";
        document.head.appendChild(st);
    })();

    function crearMultiselect(contenedor, textoTodos) {

        contenedor.classList.add("multiselect");

        contenedor.innerHTML = `
            <button type="button" class="ms-boton" aria-haspopup="listbox" aria-expanded="false">
                <span class="ms-texto"></span>
                <span class="ms-flecha">▾</span>
            </button>

            <div class="ms-panel" hidden>
                <input type="search" class="ms-buscar" placeholder="Buscar...">

                <div class="ms-acciones">
                    <button type="button" class="ms-todos">Seleccionar visibles</button>
                    <button type="button" class="ms-ninguno">Quitar selección</button>
                </div>

                <div class="ms-lista" role="listbox" aria-multiselectable="true"></div>
            </div>
        `;

        const boton   = contenedor.querySelector(".ms-boton");
        const texto   = contenedor.querySelector(".ms-texto");
        const panel   = contenedor.querySelector(".ms-panel");
        const buscar  = contenedor.querySelector(".ms-buscar");
        const lista   = contenedor.querySelector(".ms-lista");
        const btnTodos   = contenedor.querySelector(".ms-todos");
        const btnNinguno = contenedor.querySelector(".ms-ninguno");

        let opciones = [];
        const seleccionados = new Set();


        function notificarCambio() {
            contenedor.dispatchEvent(
                new CustomEvent("change", { bubbles: true })
            );
        }

        function actualizarTexto() {

            if (seleccionados.size === 0) {
                texto.textContent = textoTodos;
            }
            else if (seleccionados.size === 1) {
                const unica = opciones.find(function (o) {
                    return seleccionados.has(o.value);
                });
                texto.textContent = unica ? unica.label : "1 seleccionado";
            }
            else {
                texto.textContent = seleccionados.size + " seleccionados";
            }
        }

        function dibujar() {

            const termino = buscar.value.trim().toLowerCase();

            lista.innerHTML = "";

            opciones.forEach(function (opcion) {

                if (
                    termino &&
                    !opcion.label.toLowerCase().includes(termino)
                ) {
                    return;
                }

                const fila = document.createElement("label");
                fila.className = "ms-opcion";

                const check = document.createElement("input");
                check.type = "checkbox";
                check.value = opcion.value;
                check.checked = seleccionados.has(opcion.value);

                check.addEventListener("change", function () {

                    if (check.checked) {
                        seleccionados.add(opcion.value);
                    } else {
                        seleccionados.delete(opcion.value);
                    }

                    actualizarTexto();
                    notificarCambio();
                });

                const etiqueta = document.createElement("span");
                etiqueta.textContent = opcion.label;

                fila.appendChild(check);
                fila.appendChild(etiqueta);
                lista.appendChild(fila);
            });

            if (!lista.children.length) {
                lista.innerHTML =
                    '<div class="ms-vacio">Sin resultados</div>';
            }
        }

        function abrir() {
            panel.hidden = false;
            boton.setAttribute("aria-expanded", "true");
            buscar.focus();
        }

        function cerrar() {
            panel.hidden = true;
            boton.setAttribute("aria-expanded", "false");
        }

        boton.addEventListener("click", function () {
            panel.hidden ? abrir() : cerrar();
        });

        // Cerrar al hacer clic fuera o con Escape
        document.addEventListener("click", function (e) {
            if (!contenedor.contains(e.target)) cerrar();
        });

        contenedor.addEventListener("keydown", function (e) {
            if (e.key === "Escape") {
                cerrar();
                boton.focus();
            }
        });

        buscar.addEventListener("input", dibujar);

        btnTodos.addEventListener("click", function () {

            const termino = buscar.value.trim().toLowerCase();

            opciones.forEach(function (opcion) {
                if (
                    !termino ||
                    opcion.label.toLowerCase().includes(termino)
                ) {
                    seleccionados.add(opcion.value);
                }
            });

            dibujar();
            actualizarTexto();
            notificarCambio();
        });

        btnNinguno.addEventListener("click", function () {
            seleccionados.clear();
            dibujar();
            actualizarTexto();
            notificarCambio();
        });

        actualizarTexto();

        return {

            // Recibe [{ value, label }, ...]
            setOpciones: function (nuevas) {

                opciones = nuevas;

                // Conservar solo lo que sigue existiendo
                const validos = new Set(
                    nuevas.map(function (o) { return o.value; })
                );

                Array.from(seleccionados).forEach(function (v) {
                    if (!validos.has(v)) seleccionados.delete(v);
                });

                buscar.value = "";
                dibujar();
                actualizarTexto();
            },

            // Devuelve un arreglo con los valores marcados
            getSeleccionados: function () {
                return Array.from(seleccionados);
            },

            limpiar: function () {
                seleccionados.clear();
                buscar.value = "";
                dibujar();
                actualizarTexto();
            }
        };
    }


    // Instancias de los filtros con selección múltiple
    const multiPDV =
        crearMultiselect(
            document.getElementById("filtroPDV"),
            "Todos los PDV"
        );

    const multiFuncionario =
        crearMultiselect(
            document.getElementById("filtroFuncionario"),
            "Todos los funcionarios"
        );


// ============================================================
// 1. CARGAR ARCHIVO EXCEL AUTOMÁTICAMENTE
// ============================================================

async function cargarExcelAutomaticamente() {

    try {

        // El Excel se carga únicamente desde la carpeta data del proyecto
        const rutasPosibles = [
            "data/reporte_ingresos.xlsx"
        ];

        let respuesta = null;

        for (const ruta of rutasPosibles) {

            try {

                const intento = await fetch(ruta);

                if (intento.ok) {
                    respuesta = intento;
                    break;
                }
            }
            catch (errorRuta) {
                // Se intenta con la siguiente ruta
            }
        }

        if (!respuesta) {

            throw new Error(
                "No se pudo encontrar el archivo reporte_ingresos.xlsx."
            );
        }

        const datos =
            new Uint8Array(
                await respuesta.arrayBuffer()
            );

        procesarLibroExcel(datos);

    }
    catch (error) {

        console.error(
            "ERROR AL CARGAR EXCEL AUTOMÁTICAMENTE:",
            error
        );

        // Sin carga manual: solo se informa el error en pantalla.
        document.getElementById("resultado").innerHTML = `
            <p>
                ❌ No se pudo cargar
                <strong>data/reporte_ingresos.xlsx</strong>.
                Verifica que el archivo exista en la carpeta
                <strong>data</strong> del proyecto y que la página
                se abra desde un servidor (no con doble clic).
            </p>
        `;
    }
}


// ============================================================
// PROCESAR EL LIBRO DE EXCEL
// ============================================================

function procesarLibroExcel(datos) {

    const libro =
        XLSX.read(datos, {
            type: "array",
            cellDates: false
        });

    if (!libro.Sheets["programacion"]) {
        throw new Error(
            "No se encontró la hoja 'programacion'."
        );
    }

    if (!libro.Sheets["ingresos"]) {
        throw new Error(
            "No se encontró la hoja 'ingresos'."
        );
    }

    const datosProgramacion =
        XLSX.utils.sheet_to_json(
            libro.Sheets["programacion"],
            { defval: "" }
        );

    const datosIngresos =
        XLSX.utils.sheet_to_json(
            libro.Sheets["ingresos"],
            { defval: "" }
        );

    resultados =
        cruzarDatos(
            datosProgramacion,
            datosIngresos
        );

    // Actualiza filtros, indicadores, alertas y tabla
    mostrarResultados(resultados);
}


// ============================================================
// CARGAR AUTOMÁTICAMENTE AL ABRIR EL DASHBOARD
// ============================================================

console.log("🔥🔥🔥 APP.JS NUEVO CARGADO 🔥🔥🔥");

cargarExcelAutomaticamente();


    // ============================================================
    // 2. NORMALIZAR TEXTO
    // ============================================================

    function normalizarTexto(valor) {

        if (
            valor === null ||
            valor === undefined
        ) {

            return "";
        }

        return String(valor)
            .trim()
            .toUpperCase();

    }


    // ============================================================
    // 3. NORMALIZAR DOCUMENTO
    // ============================================================

    function normalizarDocumento(valor) {

        if (
            valor === null ||
            valor === undefined
        ) {

            return "";
        }

        return String(valor)
            .replace(/\D/g, "")
            .trim();

    }


    // ============================================================
    // 4. NORMALIZAR PDV
    // ============================================================

    function normalizarPDV(valor) {

        if (
            valor === null ||
            valor === undefined
        ) {

            return "";
        }

        return String(valor)
            .trim()
            .replace(/\.0$/, "");

    }


    // ============================================================
    // 5. OBTENER FECHA
    // ============================================================

    function obtenerFecha(valor) {

        if (
            valor === null ||
            valor === undefined ||
            valor === ""
        ) {

            return null;
        }


        // --------------------------------------------------------
        // Excel almacena las fechas como números.
        // --------------------------------------------------------

        if (typeof valor === "number") {

            try {

                const fechaExcel =
                    XLSX.SSF.parse_date_code(valor);

                if (!fechaExcel) {

                    return null;
                }

                return new Date(
                    fechaExcel.y,
                    fechaExcel.m - 1,
                    fechaExcel.d
                );

            }
            catch (error) {

                return null;
            }
        }


        // --------------------------------------------------------
        // Si ya es Date
        // --------------------------------------------------------

        if (valor instanceof Date) {

            return new Date(
                valor.getFullYear(),
                valor.getMonth(),
                valor.getDate()
            );
        }


        // --------------------------------------------------------
        // Si viene como texto
        // --------------------------------------------------------

        const texto =
            String(valor).trim();


        // DD/MM/YYYY

        if (texto.includes("/")) {

            const partes =
                texto.split("/");

            if (partes.length >= 3) {

                const dia =
                    parseInt(partes[0], 10);

                const mes =
                    parseInt(partes[1], 10);

                const anio =
                    parseInt(
                        partes[2].split(" ")[0],
                        10
                    );

                if (
                    !isNaN(dia) &&
                    !isNaN(mes) &&
                    !isNaN(anio)
                ) {

                    return new Date(
                        anio,
                        mes - 1,
                        dia
                    );
                }
            }
        }


        // YYYY-MM-DD

        if (texto.includes("-")) {

            const partes =
                texto.split("-");

            if (partes.length >= 3) {

                const anio =
                    parseInt(partes[0], 10);

                const mes =
                    parseInt(partes[1], 10);

                const dia =
                    parseInt(
                        partes[2].split(" ")[0],
                        10
                    );

                if (
                    !isNaN(anio) &&
                    !isNaN(mes) &&
                    !isNaN(dia)
                ) {

                    return new Date(
                        anio,
                        mes - 1,
                        dia
                    );
                }
            }
        }


        return null;
    }


    // ============================================================
    // 6. OBTENER CLAVE DE FECHA
    // ============================================================

    function obtenerClaveFecha(valor) {

        const fecha =
            obtenerFecha(valor);

        if (!fecha) {

            return "";
        }

        const anio =
            fecha.getFullYear();

        const mes =
            String(
                fecha.getMonth() + 1
            ).padStart(2, "0");

        const dia =
            String(
                fecha.getDate()
            ).padStart(2, "0");


        return `${anio}-${mes}-${dia}`;
    }


    // ============================================================
    // 7. OBTENER MINUTOS DE UNA HORA
    // ============================================================

    function obtenerMinutosHora(valor) {

        if (
            valor === null ||
            valor === undefined ||
            valor === ""
        ) {
            return null;
        }


        // ========================================================
        // 1. EXCEL GUARDA FECHA/HORA COMO NÚMERO
        // ========================================================

        if (typeof valor === "number") {

            const fraccionDia = valor % 1;

            return Math.round(
                fraccionDia * 24 * 60
            );
        }


        // ========================================================
        // 2. SI YA ES UN OBJETO DATE
        // ========================================================

        if (valor instanceof Date) {

            return (
                valor.getHours() * 60 +
                valor.getMinutes()
            );
        }


        // ========================================================
        // 3. SI VIENE COMO TEXTO
        // ========================================================

        let texto = String(valor).trim();

        if (!texto) {
            return null;
        }


        // ========================================================
        // BUSCAR UNA HORA DENTRO DEL TEXTO
        //
        // Ejemplos que puede recibir:
        //
        // 22/09/2026 08:28
        // 22/09/2026 08:28:15
        // 2026-09-22 08:28
        // 2026-09-22T08:28:00
        // 08:28
        // 08:28:00
        // ========================================================

        const coincidenciaHora =
            texto.match(
                /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i
            );


        if (coincidenciaHora) {

            let horas =
                parseInt(
                    coincidenciaHora[1],
                    10
                );

            const minutos =
                parseInt(
                    coincidenciaHora[2],
                    10
                );

            const periodo =
                coincidenciaHora[4]
                    ? coincidenciaHora[4].toUpperCase()
                    : "";


            // Validar minutos

            if (
                minutos < 0 ||
                minutos > 59
            ) {
                return null;
            }


            // AM / PM

            if (
                periodo === "PM" &&
                horas < 12
            ) {
                horas += 12;
            }


            if (
                periodo === "AM" &&
                horas === 12
            ) {
                horas = 0;
            }


            // Validar hora

            if (
                horas < 0 ||
                horas > 23
            ) {
                return null;
            }


            return (
                horas * 60 +
                minutos
            );
        }


        // ========================================================
        // 4. SI NO ENCONTRÓ UNA HORA
        // ========================================================

        return null;
    }

    function minutosAHora(minutos) {

        if (minutos === null || minutos === undefined) {
            return "";
        }

        const horas =
            Math.floor(minutos / 60);

        const minutosRestantes =
            minutos % 60;

        return (
            String(horas).padStart(2, "0") +
            ":" +
            String(minutosRestantes).padStart(2, "0")
        );
    }


    // ============================================================
    // 8. FORMATEAR HORA
    // ============================================================

    function formatearHora(minutos) {

        if (
            minutos === null ||
            minutos === undefined
        ) {

            return "--:--";
        }


        minutos =
            Math.round(minutos);


        const horas =
            Math.floor(minutos / 60);

        const mins =
            minutos % 60;


        return (
            String(horas).padStart(2, "0") +
            ":" +
            String(mins).padStart(2, "0")
        );
    }


    // ============================================================
    // 9. FORMATEAR FECHA
    // ============================================================

    function formatearFecha(fecha) {

        if (!fecha) {

            return "-";
        }


        const dia =
            String(
                fecha.getDate()
            ).padStart(2, "0");


        const mes =
            String(
                fecha.getMonth() + 1
            ).padStart(2, "0");


        const anio =
            fecha.getFullYear();


        return `${dia}/${mes}/${anio}`;
    }


    // ============================================================
    // 10. FORMATEAR DIFERENCIA
    // ============================================================

    function formatearDiferencia(
        diferencia
    ) {

        if (
            diferencia === null ||
            diferencia === undefined
        ) {

            return "--";
        }


        if (diferencia < 0) {

            return `${Math.abs(diferencia)} min antes`;
        }


        if (diferencia === 0) {

            return "0 min";
        }


        return `+${diferencia} min`;
    }


    // ============================================================
    // 11. BUSCAR COLUMNA DE HORA PROGRAMADA
    // ============================================================
    //
    // Como todavía estamos verificando el nombre exacto de la
    // columna en tu Excel, buscamos varias posibilidades.
    // ============================================================

    function obtenerColumnaHoraProgramada(
        registro
    ) {

        const posibles = [

            "Hora Entrada",

            "Hora de Entrada",

            "Hora Entrada Programada",

            "Hora de Entrada Programada",

            "Hora Ingreso",

            "Hora Programada",

            "Hora inicio",

            "Hora Inicio"

        ];


        for (
            const nombre of posibles
        ) {

            if (
                Object.prototype.hasOwnProperty.call(
                    registro,
                    nombre
                )
            ) {

                return registro[nombre];
            }
        }


        // Búsqueda flexible

        const columnas =
            Object.keys(registro);


        const columnaEncontrada =
            columnas.find(function (columna) {

                const texto =
                    normalizarTexto(columna);

                return (
                    texto.includes("HORA") &&
                    (
                        texto.includes("ENTRADA") ||
                        texto.includes("INGRESO") ||
                        texto.includes("INICIO")
                    )
                );

            });


        if (columnaEncontrada) {

            return registro[columnaEncontrada];
        }


        return null;
    }


    // ============================================================
    // 12. CREAR CLAVE DE CRUCE
    // ============================================================

    function crearClave(
        documento,
        pdv,
        fecha
    ) {

        return (
            normalizarDocumento(documento) +
            "|" +
            normalizarPDV(pdv) +
            "|" +
            obtenerClaveFecha(fecha)
        );
    }


    // ============================================================
    // 13. CRUZAR DATOS
    // ============================================================

    function cruzarDatos(programacion,ingresos) {

        const mapaIngresos =
            new Map();


        // --------------------------------------------------------
        // INDEXAR INGRESOS
        // --------------------------------------------------------

        ingresos.forEach(function (ingreso) {

            const documento =
                ingreso["Doc. Identidad"];

            const pdv =
                ingreso["Id Pdv"];

            const fechaRegistro =
                ingreso["Fecha registro"];


            const fecha =
                obtenerFecha(
                    fechaRegistro
                );


            if (!fecha) {

                return;
            }


            const minutosIngreso =
                obtenerMinutosHora(
                    fechaRegistro
                );


            if (
                minutosIngreso === null
            ) {

                return;
            }


            const clave =
                crearClave(
                    documento,
                    pdv,
                    fechaRegistro
                );


            // Si existen varios registros para el mismo
            // empleado + PDV + fecha, tomamos el primero.

            if (!mapaIngresos.has(clave)) {

                mapaIngresos.set(
                    clave,
                    {
                        documento:
                            normalizarDocumento(
                                documento
                            ),

                        pdv:
                            normalizarPDV(
                                pdv
                            ),

                        fecha: fecha,

                        minutosIngreso:
                            minutosIngreso,

                        registroOriginal:
                            ingreso
                    }
                );

            }
            else {

                const existente =
                    mapaIngresos.get(clave);


                // Conservamos el ingreso más temprano.

                if (
                    minutosIngreso <
                    existente.minutosIngreso
                ) {

                    mapaIngresos.set(
                        clave,
                        {
                            documento:
                                normalizarDocumento(
                                    documento
                                ),

                            pdv:
                                normalizarPDV(
                                    pdv
                                ),

                            fecha: fecha,

                            minutosIngreso:
                                minutosIngreso,

                            registroOriginal:
                                ingreso
                        }
                    );
                }
            }

        });


        console.log(
            "Mapa de ingresos:",
            mapaIngresos
        );


        // --------------------------------------------------------
        // RECORRER PROGRAMACIÓN
        // --------------------------------------------------------

        const resultado = [];


        programacion.forEach(
            function (programado) {

                const documento =
                    programado["Nro Doc Ident"];

                const pdv =
                    programado["ID PDV"];

                const fechaProgramacion =
                    programado[
                        "Fecha de Programación"
                    ];


                // Buscar hora programada
                const horaProgramada =
                    obtenerColumnaHoraProgramada(
                        programado
                    );


                const fecha =
                    obtenerFecha(
                        fechaProgramacion
                    );


                const minutosProgramados =
                    obtenerMinutosHora(
                        horaProgramada
                    );


                const clave =
                    crearClave(
                        documento,
                        pdv,
                        fechaProgramacion
                    );


                const ingreso =
                    mapaIngresos.get(
                        clave
                    );


                const fila = {
                    documento:
                        normalizarDocumento(
                            documento
                        ),

                    pdv:
                        normalizarPDV(
                            pdv
                        ),

                    nombrePDV:
                        programado["Nombre alias"] || "",

                    funcionario:
                        programado["Funcionario"] || "",

                    fecha:
                        fecha,

                    horaProgramada:
                        minutosProgramados,

                    horaReal:
                        null,

                    diferencia:
                        null,

                    estado:
                        "AÚN NO MARCA",

                    // ========================================================
                    // NUEVOS DATOS DE INGRESOS
                    // ========================================================

                    fechaSalida:
                        ingreso
                            ? ingreso.registroOriginal["Fecha Registro Salida"]
                            : "",

                    tiempoCumplido:
                        ingreso
                            ? ingreso.registroOriginal["Tiempo Cumplido"]
                            : "",

                    imagenIngreso:
                        ingreso
                            ? ingreso.registroOriginal["Imagen Ingreso"]
                            : "",

                    imagenSalida:
                        ingreso
                            ? ingreso.registroOriginal["Imagen Salida"]
                            : "",

                    idPdv:
                        ingreso
                            ? ingreso.registroOriginal["Id Pdv"]
                            : ""
                };


                // ------------------------------------------------
                // SI NO MARCÓ
                // ------------------------------------------------

                if (!ingreso) {

                    resultado.push(fila);

                    return;
                }


                // ------------------------------------------------
                // SI MARCÓ
                // ------------------------------------------------

                fila.horaReal =
                    ingreso.minutosIngreso;


                if (
                    minutosProgramados !== null
                ) {

                    fila.diferencia =
                        ingreso.minutosIngreso -
                        minutosProgramados;


                    if (
                        fila.diferencia < 0
                    ) {

                        fila.estado =
                            "TEMPRANO";

                    }
                    else if (
                        fila.diferencia === 0
                    ) {

                        fila.estado =
                            "A TIEMPO";

                    }
                    else {

                        fila.estado =
                            "TARDE";
                    }

                }


                resultado.push(fila);

            }
        );


        return resultado;
    }


    // ============================================================
    // 14. MOSTRAR RESULTADOS
    // ============================================================

    function mostrarResultados(datos, actualizarPDV = true) {

    console.log("🔥 NUEVA VERSION mostrarResultados EJECUTADA");

    const contenedor =
        document.getElementById("resultado");

    if (!contenedor) {
        return;
    }

    if (actualizarPDV) {
        cargarFiltroPDV();
        cargarFiltroFuncionario();
    }

    actualizarAlertasRecurrencia(datos);

    window.mostrarFotos = function (fila) {

    const imagenIngreso = fila.imagenIngreso;
    const imagenSalida = fila.imagenSalida;

    let contenido = "";

    if (imagenIngreso) {
        contenido += `
            <div>
                <h3>Foto ingreso</h3>
                <img
                    src="${imagenIngreso}"
                    alt="Foto ingreso"
                    style="
                        max-width: 100%;
                        max-height: 400px;
                        border-radius: 8px;
                    "
                >
            </div>
        `;
    }

    if (imagenSalida) {
        contenido += `
            <div>
                <h3>Foto salida</h3>
                <img
                    src="${imagenSalida}"
                    alt="Foto salida"
                    style="
                        max-width: 100%;
                        max-height: 400px;
                        border-radius: 8px;
                    "
                >
            </div>
        `;
    }

    if (!contenido) {
        alert("No hay fotos disponibles para este registro.");
        return;
    }

    const ventana = window.open(
        "",
        "_blank",
        "width=900,height=700"
    );

    ventana.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Fotos de ingreso y salida</title>
        </head>

        <body style="
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
        ">

            <h2>Fotos del registro</h2>

            <div style="
                display: flex;
                gap: 30px;
                flex-wrap: wrap;
            ">
                ${contenido}
            </div>

        </body>
        </html>
    `);

    ventana.document.close();
}


    function cargarFiltroFuncionario() {


        const funcionarios = [
            ...new Set(
                resultados
                    .map(
                        fila => fila.funcionario
                    )
                    .filter(
                        funcionario =>
                            funcionario !== ""
                    )
            )
        ];

        funcionarios.sort((a, b) =>
            String(a).localeCompare(
                String(b),
                undefined,
                { numeric: true }
            )
        );

        multiFuncionario.setOpciones(
            funcionarios.map(function (funcionario) {
                return {
                    value: String(funcionario),
                    label: String(funcionario)
                };
            })
        );
    }


    // --------------------------------------------------------
    // CALCULAR INDICADORES
    // --------------------------------------------------------

    const totalProgramados =
        datos.length;

    const totalIngresos =
        datos.filter(
            function (fila) {
                return fila.horaReal !== null;
            }
        ).length;

    const totalTarde =
        datos.filter(
            function (fila) {
                return fila.estado === "TARDE";
            }
        ).length;

    const totalPendientes =
        datos.filter(
            function (fila) {
                return (
                    fila.estado ===
                    "AÚN NO MARCA"
                );
            }
        ).length;


    // --------------------------------------------------------
    // ACTUALIZAR TARJETAS
    // --------------------------------------------------------

    document.getElementById(
        "totalProgramados"
    ).textContent =
        totalProgramados;

    document.getElementById(
        "totalIngresos"
    ).textContent =
        totalIngresos;

    document.getElementById(
        "totalTarde"
    ).textContent =
        totalTarde;

    document.getElementById(
        "totalPendientes"
    ).textContent =
        totalPendientes;


    // --------------------------------------------------------
    // SI NO HAY RESULTADOS
    // --------------------------------------------------------

    if (datos.length === 0) {

        contenedor.innerHTML = `
            <p>
                No se encontraron registros
                de programación.
            </p>
        `;

        return;
    }


    // --------------------------------------------------------
    // TABLA
    // --------------------------------------------------------

    let html = `

        <div class="tabla-contenedor">

            <table>

                <thead>

                    <tr>

                        <th>Funcionario</th>

                        <th>PDV</th>

                        <th>Fecha</th>

                        <th>Hora programada</th>

                        <th>Hora ingreso</th>

                        <th>Hora salida</th>

                        <th>Tiempo cumplido</th>

                        <th>Diferencia</th>

                        <th>Estado</th>

                        <th>Fotos</th>

                    </tr>

                </thead>

                <tbody>

    `;


    datos.forEach(
        function (fila) {

            let clase = "";
            let icono = "";


            if (
                fila.estado ===
                "TEMPRANO"
            ) {

                clase = "temprano";
                icono = "🟢";

            }
            else if (
                fila.estado ===
                "A TIEMPO"
            ) {

                clase = "a-tiempo";
                icono = "🔵";

            }
            else if (
                fila.estado ===
                "TARDE"
            ) {

                clase = "tarde";
                icono = "🔴";

            }
            else {

                clase = "pendiente";
                icono = "🟡";

            }


            // ------------------------------------------------
            // HORA DE SALIDA
            // ------------------------------------------------

            let horaSalida = "-";

            if (fila.fechaSalida) {

                horaSalida =
                    String(formatearFechaHoraSalida(
                        fila.fechaSalida
                    )).slice(0, 5);

            }


            // ------------------------------------------------
            // TIEMPO CUMPLIDO
            // ------------------------------------------------

            let tiempoCumplido = "-";

            if (fila.tiempoCumplido) {

                tiempoCumplido =
                    formatearTiempoCumplido(
                        fila.tiempoCumplido
                    );

            }


            // ------------------------------------------------
            // FOTOS
            // ------------------------------------------------

            let botonFotos = "-";

            if (
                fila.imagenIngreso ||
                fila.imagenSalida
            ) {

                botonFotos = `

                    <button
                        class="btn-fotos"
                        onclick="mostrarFotos(
                            ${JSON.stringify(fila).replace(/"/g, '&quot;')}
                        )"
                    >
                        📷 Ver fotos
                    </button>

                `;

            }


            html += `

                <tr>

                    <td>
                        ${fila.funcionario || "-"}
                    </td>

                    <td>
                        ${fila.nombrePDV || "-"}
                    </td>

                    <td>
                        ${formatearFecha(
                            fila.fecha
                        )}
                    </td>

                    <td>
                        ${formatearHora(
                            fila.horaProgramada
                        )}
                    </td>

                    <td>
                        ${formatearHora(
                            fila.horaReal
                        )}
                    </td>

                    <td>
                        ${horaSalida}
                    </td>

                    <td>
                        <strong>
                            ${tiempoCumplido}
                        </strong>
                    </td>

                    <td>
                        ${formatearDiferencia(
                            fila.diferencia
                        )}
                    </td>

                    <td>

                        <span class="estado ${clase}">

                            ${icono}

                            ${fila.estado}

                        </span>

                    </td>

                    <td>
                        ${botonFotos}
                    </td>

                </tr>

            `;

        }
    );


    html += `

                </tbody>

            </table>

        </div>

    `;


    contenedor.innerHTML =
        html;

}

function formatearFechaHoraSalida(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return "-";
    }

    function dosDigitos(n) {
        return String(n).padStart(2, "0");
    }

    // Excel entrega fecha/hora como número (fracción del día)
    if (typeof valor === "number") {

        const fraccion = valor - Math.floor(valor);
        let totalMinutos = Math.round(fraccion * 24 * 60);

        if (totalMinutos >= 1440) {
            totalMinutos = totalMinutos - 1440;
        }

        const horas = Math.floor(totalMinutos / 60);
        const minutos = totalMinutos % 60;

        return dosDigitos(horas) + ":" + dosDigitos(minutos);
    }

    if (valor instanceof Date && !isNaN(valor.getTime())) {
        return dosDigitos(valor.getHours()) + ":" + dosDigitos(valor.getMinutes());
    }

    const texto = String(valor).trim();

    // Buscar la hora dentro del texto (ej: "2024-05-10 14:35:00")
    const coincidencia = texto.match(/(\d{1,2}):(\d{2})/);

    if (coincidencia) {

        let horas = parseInt(coincidencia[1], 10);
        const minutos = coincidencia[2];

        if (/p\.?\s?m\.?/i.test(texto) && horas < 12) {
            horas = horas + 12;
        }

        if (/a\.?\s?m\.?/i.test(texto) && horas === 12) {
            horas = 0;
        }

        return dosDigitos(horas) + ":" + minutos;
    }

    const fecha = new Date(texto);

    if (!isNaN(fecha.getTime())) {
        return dosDigitos(fecha.getHours()) + ":" + dosDigitos(fecha.getMinutes());
    }

    return "-";
}

function formatearTiempoCumplido(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return "En curso";
    }

    // Excel puede entregar la duración
    // como una fracción de día.
    if (typeof valor === "number") {

        const totalMinutos =
            Math.round(
                valor * 24 * 60
            );

        const horas =
            Math.floor(
                totalMinutos / 60
            );

        const minutos =
            totalMinutos % 60;

        return `${horas}h ${String(minutos).padStart(2, "0")}m`;
    }

    return valor;
}

    // ============================================================
    // FILTROS
    // ============================================================


    // ------------------------------------------------------------
    // 1. LLENAR LISTA DE PDV
    // ------------------------------------------------------------

    function cargarFiltroPDV() {



        // Obtener código y nombre de cada PDV
        const pdvs = [
            ...new Map(
                resultados
                    .filter(
                        fila => fila.pdv !== ""
                    )
                    .map(fila => [
                        fila.pdv,
                        fila.nombrePDV
                    ])
            ).entries()
        ];


        // Ordenar los PDV por nombre
        pdvs.sort((a, b) =>
            String(a[1]).localeCompare(
                String(b[1]),
                undefined,
                { numeric: true }
            )
        );


        // Crear las opciones (internamente se usa el código,
        // el usuario ve el nombre)
        multiPDV.setOpciones(
            pdvs.map(function ([codigo, nombre]) {
                return {
                    value: String(codigo),
                    label: nombre || String(codigo)
                };
            })
        );
    }


    // ------------------------------------------------------------
    // 2. APLICAR FILTROS
    // ------------------------------------------------------------

    function aplicarFiltros() {

        const filtroFechaInicio =
            document.getElementById("filtroFechaInicio").value;

        const filtroFechaFin =
            document.getElementById("filtroFechaFin").value;


        const filtroPDV = multiPDV.getSeleccionados();


        const filtroEstado =
            document.getElementById(
                "filtroEstado"
            ).value;


        const filtroFuncionario = multiFuncionario.getSeleccionados();


        // --------------------------------------------------------
        // Filtrar resultados
        // --------------------------------------------------------

        const filtrados =
            resultados.filter(function (fila) {


                // -----------------------------------------------
                // FILTRO FECHA
                // -----------------------------------------------

                if (filtroFechaInicio) {

        if (
            fila.fecha &&
            obtenerClaveFecha(fila.fecha) <
                filtroFechaInicio
                ) {
                    return false;
                }
            }


            if (filtroFechaFin) {

                if (
                    fila.fecha &&
                    obtenerClaveFecha(fila.fecha) >
                        filtroFechaFin
                ) {
                    return false;
                }
            }


                // -----------------------------------------------
                // FILTRO PDV
                // -----------------------------------------------

                if (
                    filtroPDV.length > 0 &&
                    !filtroPDV.includes(String(fila.pdv))
                ) {
                    return false;
                }


                // -----------------------------------------------
                // FILTRO ESTADO
                // -----------------------------------------------

                if (filtroEstado) {

                    if (
                        fila.estado !==
                        filtroEstado
                    ) {

                        return false;
                    }

                }


                // -----------------------------------------------
                // FILTRO DOCUMENTO
                // -----------------------------------------------

                if (
                    filtroFuncionario.length > 0 &&
                    !filtroFuncionario.includes(String(fila.funcionario))
                ) {
                    return false;
                }


                return true;

            });


        // Mostrar resultados filtrados
        mostrarResultados(
            filtrados,
            false
        );

    }


    // ------------------------------------------------------------
    // 3. EVENTOS DE LOS FILTROS
    // ------------------------------------------------------------

    document.getElementById("filtroFechaInicio")
        .addEventListener(
            "change",
            aplicarFiltros
        );


    document.getElementById("filtroFechaFin")
        .addEventListener(
            "change",
            aplicarFiltros
        );


    document
        .getElementById("filtroPDV")
        .addEventListener(
            "change",
            aplicarFiltros
        );


    document
        .getElementById("filtroEstado")
        .addEventListener(
            "change",
            aplicarFiltros
        );


    document.getElementById("filtroFuncionario")
        .addEventListener(
            "change",
            aplicarFiltros
        );


    // ------------------------------------------------------------
    // 4. LIMPIAR FILTROS
    // ------------------------------------------------------------

    document
        .getElementById("btnLimpiarFiltros")
        .addEventListener(
            "click",
            function () {

                document.getElementById(
                "filtroFechaInicio"
                ).value = "";

                document.getElementById(
                "filtroFechaFin"
                ).value = "";


                multiPDV.limpiar();


                document.getElementById(
                    "filtroEstado"
                ).value = "";


                multiFuncionario.limpiar();


                mostrarResultados(
                    resultados,
                    false
                );

            }
        );
        function exportarExcel() {

        // Obtener los filtros actuales
        const filtroFechaInicio =
            document.getElementById("filtroFechaInicio").value;

        const filtroFechaFin =
            document.getElementById("filtroFechaFin").value;

        const filtroPDV = multiPDV.getSeleccionados();

        const filtroEstado =
            document.getElementById("filtroEstado").value;

        const filtroFuncionario = multiFuncionario.getSeleccionados();


        // Aplicar los mismos filtros de la pantalla
        const datosExportar =
            resultados.filter(function (fila) {

                if (filtroFechaInicio) {

                    if (
                        fila.fecha &&
                        obtenerClaveFecha(fila.fecha) <
                            filtroFechaInicio
                    ) {
                        return false;
                    }
                }


                if (filtroFechaFin) {

                    if (
                        fila.fecha &&
                        obtenerClaveFecha(fila.fecha) >
                            filtroFechaFin
                    ) {
                        return false;
                    }
                }


                if (
                    filtroPDV.length > 0 &&
                    !filtroPDV.includes(String(fila.pdv))
                ) {
                    return false;
                }


                if (filtroEstado) {

                    if (fila.estado !== filtroEstado) {
                        return false;
                    }
                }


                if (
                    filtroFuncionario.length > 0 &&
                    !filtroFuncionario.includes(String(fila.funcionario))
                ) {
                    return false;
                }


                return true;
            });


        // Validar que existan datos
        if (datosExportar.length === 0) {

            alert(
                "No hay datos para exportar con los filtros seleccionados."
            );

            return;
        }


        // Preparar información para Excel
        const datosExcel =
            datosExportar.map(function (fila) {

                return {

                    "Funcionario":
                        fila.funcionario || "",

                    "PDV":
                        fila.nombrePDV || "",

                    "Fecha":
                        obtenerClaveFecha(fila.fecha),

                    "Hora programada":
                        fila.horaProgramada !== null
                            ? minutosAHora(fila.horaProgramada)
                            : "",

                    "Hora real":
                        fila.horaReal !== null
                            ? minutosAHora(fila.horaReal)
                            : "",

                    "Diferencia minutos":
                        fila.diferencia !== null
                            ? fila.diferencia
                            : "",

                    "Estado":
                        fila.estado || ""
                };
            });


        // Crear archivo Excel
        const hoja =
            XLSX.utils.json_to_sheet(datosExcel);

        const libro =
            XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            libro,
            hoja,
            "Detalle"
        );


        // Nombre del archivo
        let nombreArchivo =
            "Reporte_Ingresos";

        if (
            filtroFechaInicio &&
            filtroFechaFin
        ) {

            nombreArchivo +=
                "_" +
                filtroFechaInicio +
                "_al_" +
                filtroFechaFin;

        } else if (filtroFechaInicio) {

            nombreArchivo +=
                "_desde_" +
                filtroFechaInicio;

        } else if (filtroFechaFin) {

            nombreArchivo +=
                "_hasta_" +
                filtroFechaFin;
        }


        nombreArchivo += ".xlsx";


        // Descargar Excel
        XLSX.writeFile(
            libro,
            nombreArchivo
        );
    }
        function actualizarAlertasRecurrencia(datos) {

        const contenedor =
            document.getElementById("alertasRecurrencia");

        if (!contenedor) return;


        // Agrupar información por funcionario
        const resumen = {};


        datos.forEach(function (fila) {

            const funcionario =
                fila.funcionario || "Sin funcionario";


            if (!resumen[funcionario]) {

                resumen[funcionario] = {
                    programados: 0,
                    ingresos: 0,
                    tardanzas: 0,
                    minutosTarde: 0,
                    ultimaTardanza: null
                };
            }


            // Día programado
            resumen[funcionario].programados++;


            // Si registró ingreso
            if (fila.horaReal !== null) {

                resumen[funcionario].ingresos++;
            }


            // Si llegó tarde
            if (fila.estado === "TARDE") {

                resumen[funcionario].tardanzas++;


                if (fila.diferencia !== null) {

                    resumen[funcionario].minutosTarde +=
                        fila.diferencia;
                }


                // Guardar la última fecha de tardanza
                if (
                    !resumen[funcionario].ultimaTardanza ||
                    obtenerClaveFecha(fila.fecha) >
                        obtenerClaveFecha(
                            resumen[funcionario].ultimaTardanza
                        )
                ) {

                    resumen[funcionario].ultimaTardanza =
                        fila.fecha;
                }
            }
        });


        // Convertir el objeto en arreglo
        const funcionarios =
            Object.entries(resumen)
                .map(function ([nombre, datosFuncionario]) {

                    const porcentaje =
                        datosFuncionario.ingresos > 0
                            ? (
                                datosFuncionario.tardanzas /
                                datosFuncionario.ingresos
                            ) * 100
                            : 0;


                    const promedio =
                        datosFuncionario.tardanzas > 0
                            ? (
                                datosFuncionario.minutosTarde /
                                datosFuncionario.tardanzas
                            )
                            : 0;


                    return {
                        funcionario: nombre,
                        programados:
                            datosFuncionario.programados,
                        ingresos:
                            datosFuncionario.ingresos,
                        tardanzas:
                            datosFuncionario.tardanzas,
                        porcentaje: porcentaje,
                        promedioMinutos: promedio,
                        ultimaTardanza:
                            datosFuncionario.ultimaTardanza
                    };
                });


        // Solo funcionarios con 3 o más tardanzas
        const alertas =
            funcionarios.filter(function (funcionario) {

                return funcionario.tardanzas >= 3;

            });


        // Ordenar de mayor a menor cantidad de tardanzas
        alertas.sort(function (a, b) {

            return b.tardanzas - a.tardanzas;

        });


        // Si no existen alertas
        if (alertas.length === 0) {

            contenedor.innerHTML = `
                <div class="sin-alertas">
                    ✓ No se identifican funcionarios con
                    3 o más llegadas tarde en el período seleccionado.
                </div>
            `;

            return;
        }


        // Construir tabla
        let html = `
            <div class="alerta-resumen">
                Se identificaron
                <strong>${alertas.length}</strong>
                funcionario(s) con 3 o más llegadas tarde
                durante el período seleccionado.
            </div>

            <div class="tabla-alertas-contenedor">

                <table class="tabla-alertas">

                    <thead>
                        <tr>
                            <th>Funcionario</th>
                            <th>Programados</th>
                            <th>Ingresos</th>
                            <th>Tardanzas</th>
                            <th>% tardanzas</th>
                            <th>Prom. minutos tarde</th>
                            <th>Última tardanza</th>
                        </tr>
                    </thead>

                    <tbody>
        `;


        alertas.forEach(function (funcionario) {

            html += `
                <tr>

                    <td>
                        ${funcionario.funcionario}
                    </td>

                    <td>
                        ${funcionario.programados}
                    </td>

                    <td>
                        ${funcionario.ingresos}
                    </td>

                    <td class="cantidad-tardanzas">
                        ${funcionario.tardanzas}
                    </td>

                    <td>
                        ${funcionario.porcentaje.toFixed(1)}%
                    </td>

                    <td>
                        ${funcionario.promedioMinutos.toFixed(1)} min
                    </td>

                    <td>
                        ${
                            funcionario.ultimaTardanza
                                ? obtenerClaveFecha(
                                    funcionario.ultimaTardanza
                                )
                                : "-"
                        }
                    </td>

                </tr>
            `;
        });


        html += `
                    </tbody>

                </table>

            </div>
        `;


        contenedor.innerHTML = html;
    }

    // ================================
    // CAMBIO DE TEMA
    // ================================

    const btnTema =
        document.getElementById("btnTema");

    btnTema.addEventListener(
        "change",
        function () {

            document.body.classList.toggle(
                "tema-oscuro",
                btnTema.checked
            );

        }
    );


    // ================================
    // EXPORTAR A EXCEL (un solo listener)
    // ================================

    document
        .getElementById("btnExportarExcel")
        .addEventListener("click", exportarExcel);
