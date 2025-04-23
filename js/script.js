/*
 * Copyright 2025 [Tu Nombre o Nombre de tu Sitio/Empresa]. Todos los derechos reservados.
 * Script para la Calculadora de Materiales Tablayeso.
 * Maneja la lógica de agregar ítems, calcular materiales y generar reportes.
 * Implementa el criterio de cálculo v2.0, con nombres específicos para Durock Calibre 20 y lógica de tornillos de 1".
 */

document.addEventListener('DOMContentLoaded', () => {
    const itemsContainer = document.getElementById('items-container');
    const addItemBtn = document.getElementById('add-item-btn');
    const calculateBtn = document.getElementById('calculate-btn');
    const resultsContent = document.getElementById('results-content');
    const downloadOptionsDiv = document.querySelector('.download-options');
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const generateExcelBtn = document.getElementById('generate-excel-btn');

    let itemCounter = 0; // To give unique IDs to item blocks

    // Variables to store the last calculated state (needed for PDF/Excel)
    // These will now store sums of rounded quantities per item
    let lastCalculatedTotalMaterials = {};
    let lastCalculatedItemsSpecs = []; // Specs of items included in calculation
    let lastErrorMessages = []; // Store errors as an array of strings


    // --- Helper Function for Rounding Up Final Units (Applies per item material quantity) ---
    const roundUpFinalUnit = (num) => Math.ceil(num);

    // --- Helper Function to get display name for item type ---
    const getItemTypeName = (typeValue) => {
        switch (typeValue) {
            case 'muro-yeso': return 'Muro de Panel de Yeso';
            case 'cielo-yeso': return 'Cielo Falso de Panel de Yeso';
            case 'muro-durock': return 'Muro de Durock';
            default: return 'Ítem Desconocido';
        }
    };

    // --- Helper Function to get the unit for a given material name ---
    const getMaterialUnit = (materialName) => {
         // Map specific names to units based on the new criterion
        switch (materialName) {
            case 'Paneles de Yeso': return 'Und';
            case 'Paneles de Durock': return 'Und'; // Added specific Durock panel name
            case 'Postes': return 'Und';
            case 'Postes Calibre 20': return 'Und'; // Added specific Durock postes name
            case 'Canales': return 'Und';
            case 'Canales Calibre 20': return 'Und'; // Added specific Durock canales name
            case 'Pasta': return 'Caja';
            case 'Cinta de Papel': return 'm';
            case 'Lija Grano 120': return 'Pliego';
            case 'Clavos con Roldana': return 'Und';
            case 'Fulminantes': return 'Und';
            case 'Tornillos de 1" punta fina': return 'Und';
            case 'Tornillos de 1/2" punta fina': return 'Und';
            case 'Canal Listón': return 'Und';
            case 'Canal Soporte': return 'Und';
            case 'Angular de Lámina': return 'Und';
            case 'Tornillos de 1" punta broca': return 'Und';
            case 'Tornillos de 1/2" punta broca': return 'Und';
            case 'Patas': return 'Und';
            case 'Canal Listón (para cuelgue)': return 'Und'; // This is the count of profiles needed for hangers
            case 'Basecoat': return 'Saco';
            case 'Cinta malla': return 'm';
            default: return 'Und'; // Default unit if not specified
        }
    };


    // --- Function to Update Input Visibility WITHIN an Item Block ---
    const updateItemInputVisibility = (itemBlock) => {
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        const heightInputGroup = itemBlock.querySelector('.item-height-input');
        const lengthInputGroup = itemBlock.querySelector('.item-length-input');
        const facesInputGroup = itemBlock.querySelector('.item-faces-input');
        const plenumInputGroup = itemBlock.querySelector('.item-plenum-input');
        const doubleStructureInputGroup = itemBlock.querySelector('.item-double-structure-input');
        const postSpacingInputGroup = itemBlock.querySelector('.item-post-spacing-input');


        const type = structureTypeSelect.value;

        // Reset visibility for inputs within this block
        heightInputGroup.classList.add('hidden');
        lengthInputGroup.classList.add('hidden');
        facesInputGroup.classList.add('hidden');
        plenumInputGroup.classList.add('hidden');
        doubleStructureInputGroup.classList.add('hidden');
        postSpacingInputGroup.classList.add('hidden');


        // Set visibility based on selected type for THIS block
        if (type === 'muro-yeso' || type === 'muro-durock') {
            heightInputGroup.classList.remove('hidden');
            facesInputGroup.classList.remove('hidden');
            doubleStructureInputGroup.classList.remove('hidden'); // Double structure applies to walls
            postSpacingInputGroup.classList.remove('hidden'); // Post spacing applies to walls
             // Hide ceiling-specific inputs
            lengthInputGroup.classList.add('hidden');
            plenumInputGroup.classList.add('hidden');

        } else if (type === 'cielo-yeso') {
            lengthInputGroup.classList.remove('hidden');
            plenumInputGroup.classList.remove('hidden');
            // Hide wall-specific inputs for ceiling
            heightInputGroup.classList.add('hidden');
            facesInputGroup.classList.add('hidden');
            doubleStructureInputGroup.classList.add('hidden');
            postSpacingInputGroup.classList.add('hidden');

        } else {
             // Hide all type-specific inputs if type is unknown (shouldn't happen with current options)
            heightInputGroup.classList.add('hidden');
            lengthInputGroup.classList.add('hidden');
            facesInputGroup.classList.add('hidden');
            plenumInputGroup.classList.add('hidden');
            doubleStructureInputGroup.classList.add('hidden');
            postSpacingInputGroup.classList.add('hidden');
        }
    };

    // --- Function to Create an Item Input Block ---
    const createItemBlock = () => {
        itemCounter++;
        const itemId = `item-${itemCounter}`;

        const itemHtml = `
            <div class="item-block" data-item-id="${itemId}">
                <h3>Muro/Cielo #${itemCounter}</h3>
                <button class="remove-item-btn">Eliminar</button>
                <div class="input-group">
                    <label for="type-${itemId}">Tipo de Estructura:</label>
                    <select class="item-structure-type" id="type-${itemId}">
                        <option value="muro-yeso">Muro de Panel de Yeso</option>
                        <option value="cielo-yeso">Cielo Falso de Panel de Yeso</option>
                        <option value="muro-durock">Muro de Durock</option>
                    </select>
                </div>

                <div class="input-group item-width-input">
                    <label for="width-${itemId}">Ancho (m):</label>
                    <input type="number" class="item-width" id="width-${itemId}" step="0.01" min="0" value="3.0">
                </div>

                <div class="input-group item-height-input">
                    <label for="height-${itemId}">Alto (m):</label>
                    <input type="number" class="item-height" id="height-${itemId}" step="0.01" min="0" value="2.4">
                </div>

                <div class="input-group item-length-input hidden">
                    <label for="length-${itemId}">Largo (m):</label>
                    <input type="number" class="item-length" id="length-${itemId}" step="0.01" min="0" value="4.0">
                </div>

                <div class="input-group item-faces-input">
                    <label for="faces-${itemId}">Nº de Caras (1 o 2):</label>
                    <input type="number" class="item-faces" id="faces-${itemId}" step="1" min="1" max="2" value="1">
                </div>

                 <div class="input-group item-post-spacing-input hidden">
                    <label for="post-spacing-${itemId}">Espaciamiento Postes (m):</label>
                    <input type="number" class="item-post-spacing" id="post-spacing-${itemId}" step="0.01" min="0.1" value="0.40">
                </div>

                <div class="input-group item-plenum-input hidden">
                    <label for="plenum-${itemId}">Pleno del Cielo (m):</label>
                    <input type="number" class="item-plenum" id="plenum-${itemId}" step="0.01" min="0" value="0.5">
                </div>

                <div class="input-group item-double-structure-input">
                    <label for="double-structure-${itemId}">Estructura Doble:</label>
                    <input type="checkbox" class="item-double-structure" id="double-structure-${itemId}">
                </div>
            </div>
        `;

        const newElement = document.createElement('div');
        newElement.innerHTML = itemHtml.trim();
        const itemBlock = newElement.firstChild; // Get the actual div element

        itemsContainer.appendChild(itemBlock);

        // Add event listener to the new select element IN THIS BLOCK
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        structureTypeSelect.addEventListener('change', () => updateItemInputVisibility(itemBlock));

        // Add event listener to the new remove button
        const removeButton = itemBlock.querySelector('.remove-item-btn');
        removeButton.addEventListener('click', () => {
            itemBlock.remove(); // Remove the block from the DOM
            // Clear results and hide download buttons after removal for immediate feedback
             resultsContent.innerHTML = '<p>Ítem eliminado. Recalcula los materiales totales.</p>';
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
             // Also reset stored data on item removal
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = [];
             // Re-evaluate if calculate button should be disabled (if no items left)
             toggleCalculateButtonState();
        });


        // Set initial visibility for the inputs in the new block
        updateItemInputVisibility(itemBlock);
        // Re-evaluate if calculate button should be enabled (since an item was added)
        toggleCalculateButtonState();

        return itemBlock; // Return the created element
    };

     // --- Function to Enable/Disable Calculate Button ---
    const toggleCalculateButtonState = () => {
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');
        calculateBtn.disabled = itemBlocks.length === 0;
    };


    // --- Main Calculation Function for ALL Items ---
    const calculateMaterials = () => {
        console.log("Iniciando cálculo de materiales...");
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');
        // These will now sum up the *rounded* quantities calculated per item
        let currentTotalMaterials = {};
        let currentCalculatedItemsSpecs = []; // Array to store specs of validly calculated items
        let currentErrorMessages = []; // Use an array to collect validation error messages

        // Clear previous results and hide download buttons initially
        resultsContent.innerHTML = '';
        downloadOptionsDiv.classList.add('hidden');

        if (itemBlocks.length === 0) {
            console.log("No hay ítems para calcular.");
            resultsContent.innerHTML = '<p style="color: orange; text-align: center; font-style: italic;">Por favor, agrega al menos un Muro o Cielo para calcular.</p>';
             // Store empty results
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = ['No hay ítems agregados para calcular.'];
            return;
        }

        console.log(`Procesando ${itemBlocks.length} ítems.`);
        // Iterate through each item block and calculate its materials
        itemBlocks.forEach(itemBlock => {
            const itemNumber = itemBlock.querySelector('h3').textContent.split('#')[1]; // Extract number like "1"
            const type = itemBlock.querySelector('.item-structure-type').value;
            const width = parseFloat(itemBlock.querySelector('.item-width').value);
            const height = parseFloat(itemBlock.querySelector('.item-height').value);

            // Use explicit checks for validation and visibility
            const lengthInput = itemBlock.querySelector('.item-length');
            const length = lengthInput && !lengthInput.closest('.hidden') ? parseFloat(lengthInput.value) : NaN;

            const facesInput = itemBlock.querySelector('.item-faces');
            const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : NaN;

            const plenumInput = itemBlock.querySelector('.item-plenum');
            const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;

            const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
            const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.closest('.hidden') ? isDoubleStructureInput.checked : false;


            const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
            const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : NaN;

            console.log(`Ítem #${itemNumber}: Tipo=${type}, Ancho=${width}, Alto=${height}, Largo=${length}, Caras=${faces}, Espaciamiento Postes=${postSpacing}, Pleno=${plenum}, Doble Estructura=${isDoubleStructure}`);

             // Basic Validation for Each Item
             let itemSpecificErrors = [];

             // Validation common to all types (width)
             if (isNaN(width) || width <= 0) itemSpecificErrors.push('Ancho inválido (debe ser > 0)');

             if (type === 'muro-yeso' || type === 'muro-durock') {
                 // Wall specific validations
                 if (isNaN(height) || height <= 0) itemSpecificErrors.push('Alto inválido (debe ser > 0)');
                 if (isNaN(faces) || (faces !== 1 && faces !== 2)) itemSpecificErrors.push('Nº Caras inválido (debe ser 1 o 2)');
                 if (isNaN(postSpacing) || postSpacing <= 0) itemSpecificErrors.push('Espaciamiento Postes inválido (debe ser > 0)');
                 // Ensure width and height are positive for area calculation (used in panel/finishing calcs)
                 if (width <= 0 || height <= 0) itemSpecificErrors.push('Ancho y Alto deben ser mayores a 0 para Muros');


             } else if (type === 'cielo-yeso') {
                 // Ceiling specific validations
                 if (isNaN(length) || length <= 0) itemSpecificErrors.push('Largo inválido (debe ser > 0)');
                 // Plenum validation only if visible and required
                 if (!itemBlock.querySelector('.item-plenum-input').classList.contains('hidden') && (isNaN(plenum) || plenum < 0)) itemSpecificErrors.push('Pleno inválido (debe ser >= 0)');
                 // Length and Width must be > 0 for ceiling area
                 if (length <= 0 || width <= 0) itemSpecificErrors.push('Largo y Ancho deben ser mayores a 0 para Cielos Falsos');

             } else {
                 // Unknown type validation (shouldn't happen with current options)
                 itemSpecificErrors.push('Tipo de estructura desconocido.');
             }

             console.log(`Ítem #${itemNumber}: Errores de validación - ${itemSpecificErrors.length}`);
            // If item has errors, add to global error list and skip calculation for this item
            if (itemSpecificErrors.length > 0) {
                 const itemTitle = itemBlock.querySelector('h3').textContent;
                 currentErrorMessages.push(`Error en ${itemTitle}: ${itemSpecificErrors.join(', ')}`);
                 console.warn(`Item inválido o incompleto: ${itemTitle}. Errores: ${itemSpecificErrors.join(', ')}. Este ítem no se incluirá en el cálculo total.`);
                 return; // Skip calculation for this invalid item in the forEach
             }

             // Store specs for valid items *before* calculation (use validated inputs)
            currentCalculatedItemsSpecs.push({
                 id: itemBlock.dataset.itemId,
                 number: itemNumber,
                 type: type,
                 width: width,
                 height: height,
                 length: length, // NaN for walls
                 faces: faces,   // NaN for ceilings
                 postSpacing: postSpacing, // NaN for ceilings
                 plenum: plenum, // NaN for walls
                 isDoubleStructure: isDoubleStructure // Only relevant for walls
            });


            // Object to hold calculated materials for THIS single item (initial and intermediate floats)
            let itemMaterialsFloat = {};

            // --- Calculation Logic for the CURRENT Item (using floats initially) ---
            if (type === 'muro-yeso' || type === 'muro-durock') {
                const area = width * height; // width and height validated > 0

                // Paneles (Yeso o Durock)
                 let panelesFloat;
                 // Condición: Si el ancho del muro es menor a 0.60 m, el resultado es 1 panel (por cara).
                 if (width > 0 && width < 0.60) {
                     panelesFloat = faces; // 1 panel por cara si ancho < 0.60m
                 } else {
                     // Cálculo normal: (Ancho * Alto) / Rendimiento * Nº Caras
                     panelesFloat = (area / 2.98) * faces;
                 }
                 // Use specific names based on type
                 itemMaterialsFloat[type === 'muro-yeso' ? 'Paneles de Yeso' : 'Paneles de Durock'] = panelesFloat;


                // Postes
                // Condición: Si el ancho del muro es menor que el espaciamiento, se necesitan 2 postes.
                let postesFloat;
                if (width > 0 && postSpacing > 0 && width < postSpacing) { // Check width > 0 and postSpacing > 0
                     postesFloat = 2;
                 } else if (width > 0 && postSpacing > 0) {
                     // Cálculo: Floor(Ancho / Espaciamiento) + 1
                    postesFloat = Math.floor(width / postSpacing) + 1;
                 } else {
                     postesFloat = 0; // Should be caught by validation, but safety check
                 }
                 // Condición: Estructura Doble duplica estructura (incluye postes)
                 if (isDoubleStructure) postesFloat *= 2;
                 // Use specific names based on type
                 itemMaterialsFloat[type === 'muro-yeso' ? 'Postes' : 'Postes Calibre 20'] = postesFloat;


                // Canales
                // Condición: Si el ancho del muro es menor a 1.82 metros, el resultado es 1 canal.
                let canalesFloat;
                 if (width > 0 && width < 1.82) {
                     canalesFloat = 1;
                 } else if (width > 0) {
                     // Cálculo: (Ancho * 2) / Largo del canal (3.05)
                    canalesFloat = (width * 2) / 3.05;
                 } else {
                     canalesFloat = 0; // Should be caught by validation
                 }
                 // Condición: Estructura Doble duplica estructura (incluye canales)
                 if (isDoubleStructure) canalesFloat *= 2;
                  // Use specific names based on type
                 itemMaterialsFloat[type === 'muro-yeso' ? 'Canales' : 'Canales Calibre 20'] = canalesFloat;


                 // Acabado y Tornillos de Panel (depende del tipo específico: Yeso o Durock)
                if (type === 'muro-yeso') {
                    // Pasta (Yeso)
                    // Cálculo: (Número de paneles * 2.98) / Rendimiento (22)
                    // Use the float panel count for the area calculation part of finishing
                    itemMaterialsFloat['Pasta'] = (panelesFloat * 2.98) / 22; // Result is in cajas

                    // Cinta de Papel
                    // Cálculo: 1 metro por metro cuadrado (Cantidad de paneles * 2.98 * 1)
                     itemMaterialsFloat['Cinta de Papel'] = (panelesFloat * 2.98); // Result is in meters

                    // Lija Grano 120
                    // Cálculo: Número de paneles / 2
                     itemMaterialsFloat['Lija Grano 120'] = panelesFloat / 2;

                    // Tornillos de 1" punta fina (para panel de Yeso)
                    // Cálculo: 40 por panel
                    itemMaterialsFloat['Tornillos de 1" punta fina'] = panelesFloat * 40; // This is correct for Muro Yeso

                } else if (type === 'muro-durock') {
                    // Basecoat (Durock)
                    // Cálculo: (Cantidad de paneles * 2.98) / Rendimiento (8)
                    itemMaterialsFloat['Basecoat'] = (panelesFloat * 2.98) / 8; // Result is in sacos

                    // Cinta malla (Durock)
                    // Cálculo: 1 metro por metro cuadrado de panel (Cantidad de paneles * 2.98 * 1)
                     itemMaterialsFloat['Cinta malla'] = (panelesFloat * 2.98); // Result is in meters

                    // Tornillos de 1" punta broca (para panel de Durock)
                    // Cálculo: 40 por panel
                    itemMaterialsFloat['Tornillos de 1" punta broca'] = panelesFloat * 40; // This is correct for Muro Durock
                }

            } else if (type === 'cielo-yeso') {
                 const area = width * length; // width and length validated > 0

                 // Paneles de Yeso (Cielo)
                 // Cálculo: (Ancho * Largo) / Rendimiento (2.98)
                 let panelesCieloFloat = area / 2.98;
                 itemMaterialsFloat['Paneles de Yeso'] = panelesCieloFloat; // Name matches muro-yeso panels


                 // Canal Listón
                 // Cálculo: (Largo / 0.40) * Ancho / 3.66
                 // Using input `length` as the "Largo" and `width` as the "Ancho" as per criterion example.
                 let canalListonFloat = (length / 0.40) * width / 3.66;
                 itemMaterialsFloat['Canal Listón'] = canalListonFloat;


                 // Canal Soporte
                 // Cálculo: (Ancho / 0.90) * Largo / 3.66
                 // Using input `width` as the "Ancho" and `length` as the "Largo" as per criterion example.
                 let canalSoporteFloat = (width / 0.90) * length / 3.66;
                 itemMaterialsFloat['Canal Soporte'] = canalSoporteFloat;


                 // Angular de Lámina
                 // Cálculo: Perímetro / Largo del angular (2.44)
                 let perimetro = (width * 2 + length * 2);
                 let angularLaminaFloat = perimetro / 2.44;
                 itemMaterialsFloat['Angular de Lámina'] = angularLaminaFloat;


                // Patas (Soportes)
                // Cálculo: Cantidad de Canal Soporte * 4
                 // Use the float amount of Canal Soporte here, will round later.
                 itemMaterialsFloat['Patas'] = canalSoporteFloat * 4;


                 // Canal Listón (para cuelgue) - Represents the number of 3.66m profiles needed to cut the hangers
                 // Cálculo: Cantidad de patas * Pleno / Largo del canal listón (3.66)
                 // Use the float amount of Patas here, will round later.
                 if (plenum > 0 && itemMaterialsFloat['Patas'] > 0 && !isNaN(plenum)) { // Only calculate if plenum is positive/valid and patas > 0
                    itemMaterialsFloat['Canal Listón (para cuelgue)'] = (itemMaterialsFloat['Patas'] * plenum) / 3.66;
                 } else {
                     itemMaterialsFloat['Canal Listón (para cuelgue)'] = 0;
                 }


                 // Tornillos 1" punta broca (para panel de Cielo)
                 // Cálculo: 40 por panel
                 itemMaterialsFloat['Tornillos de 1" punta broca'] = panelesCieloFloat * 40; // This is correct for Cielo Yeso

                 // Acabado (Pasta, Cinta de Papel, Lija) - Same calculation as Muros, based on panel area
                 // Cálculo Pasta: (Cantidad de paneles * 2.98) / Rendimiento (22)
                 itemMaterialsFloat['Pasta'] = (panelesCieloFloat * 2.98) / 22;

                 // Cálculo Cinta de Papel: (Cantidad de paneles * 2.98) * 1
                 itemMaterialsFloat['Cinta de Papel'] = (panelesCieloFloat * 2.98);

                 // Cálculo Lija Grano 120: Cantidad de paneles / 2
                 itemMaterialsFloat['Lija Grano 120'] = panelesCieloFloat / 2;

            }

            // --- Round Up Quantities for Key Components to Calculate Dependent Materials ---
            // These rounded values are used *only* for calculating associated tornillos/clavos for THIS item.
            let roundedCanales = 0;
            let roundedPostes = 0;
            let roundedAngularLamina = 0;
            let roundedCanalSoporte = 0;
            let roundedCanalListon = 0;
            let roundedPatas = 0;

             // Check existence and validity before rounding - Use specific names here
            if (itemMaterialsFloat[type === 'muro-yeso' ? 'Canales' : 'Canales Calibre 20'] !== undefined && !isNaN(itemMaterialsFloat[type === 'muro-yeso' ? 'Canales' : 'Canales Calibre 20']) && itemMaterialsFloat[type === 'muro-yes' ? 'Canales' : 'Canales Calibre 20'] >= 0) {
                 roundedCanales = roundUpFinalUnit(itemMaterialsFloat[type === 'muro-yeso' ? 'Canales' : 'Canales Calibre 20']);
            }

             if (itemMaterialsFloat[type === 'muro-yeso' ? 'Postes' : 'Postes Calibre 20'] !== undefined && !isNaN(itemMaterialsFloat[type === 'muro-yeso' ? 'Postes' : 'Postes Calibre 20']) && itemMaterialsFloat[type === 'muro-yeso' ? 'Postes' : 'Postes Calibre 20'] >= 0) {
                 roundedPostes = roundUpFinalUnit(itemMaterialsFloat[type === 'muro-yeso' ? 'Postes' : 'Postes Calibre 20']);
             }

             // Cielo components always use the same names
             if (itemMaterialsFloat['Angular de Lámina'] !== undefined && !isNaN(itemMaterialsFloat['Angular de Lámina']) && itemMaterialsFloat['Angular de Lámina'] >= 0) roundedAngularLamina = roundUpFinalUnit(itemMaterialsFloat['Angular de Lámina']);
             if (itemMaterialsFloat['Canal Soporte'] !== undefined && !isNaN(itemMaterialsFloat['Canal Soporte']) && itemMaterialsFloat['Canal Soporte'] >= 0) roundedCanalSoporte = roundUpFinalUnit(itemMaterialsFloat['Canal Soporte']);
             if (itemMaterialsFloat['Canal Listón'] !== undefined && !isNaN(itemMaterialsFloat['Canal Listón']) && itemMaterialsFloat['Canal Listón'] >= 0) roundedCanalListon = roundUpFinalUnit(itemMaterialsFloat['Canal Listón']);
             // Patas calculation depends on Canal Soporte, round its calculated float value
             if (itemMaterialsFloat['Patas'] !== undefined && !isNaN(itemMaterialsFloat['Patas']) && itemMaterialsFloat['Patas'] >= 0) roundedPatas = roundUpFinalUnit(itemMaterialsFloat['Patas']);


            // --- Calculate Tornillos/Clavos Based on Rounded Component Counts for THIS Item ---
             if (type === 'muro-yeso' || type === 'muro-durock') {
                 // Clavos con Roldana (8 per Canal) - Use rounded Canales
                 itemMaterialsFloat['Clavos con Roldana'] = roundedCanales * 8;
                 itemMaterialsFloat['Fulminantes'] = itemMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                 // Tornillos 1/2" (4 per Poste) - Use rounded Postes
                 itemMaterialsFloat[type === 'muro-yeso' ? 'Tornillos de 1/2" punta fina' : 'Tornillos de 1/2" punta broca'] = roundedPostes * 4;

            } else if (type === 'cielo-yeso') {
                 // Clavos con Roldana (5 per Angular + 8 per Canal Soporte) - Use rounded counts
                 itemMaterialsFloat['Clavos con Roldana'] = (roundedAngularLamina * 5) + (roundedCanalSoporte * 8);
                 itemMaterialsFloat['Fulminantes'] = itemMaterialsFloat['Clavos con Roldana']; // Igual cantidad

                 // Tornillos 1/2" punta fina (12 per Canal Listón + 2 per Pata) - Use rounded counts
                 itemMaterialsFloat['Tornillos de 1/2" punta fina'] = (roundedCanalListon * 12) + (roundedPatas * 2);
            }

             console.log(`Ítem #${itemNumber}: Materiales calculados (float) incluyendo tornillos/clavos:`, itemMaterialsFloat);

             // --- Round Up *all* Material Quantities for THIS Item and Sum to Total ---
            let roundedItemMaterials = {}; // Store rounded quantities for this item explicitly if needed later (currently just summed)
            for (const material in itemMaterialsFloat) {
                if (itemMaterialsFloat.hasOwnProperty(material)) {
                    const floatQuantity = itemMaterialsFloat[material];
                    // Ensure the value is a valid number before rounding and summing
                    if (!isNaN(floatQuantity) && floatQuantity > 0) { // Only sum positive quantities
                         const roundedQuantity = roundUpFinalUnit(floatQuantity);
                         // Sum this rounded quantity from the current item to the overall total
                         currentTotalMaterials[material] = (currentTotalMaterials[material] || 0) + roundedQuantity;
                         roundedItemMaterials[material] = roundedQuantity; // Store the rounded value for this item
                    } else {
                         // If quantity is 0, NaN, or negative, treat as 0 for this item and don't add to total
                         roundedItemMaterials[material] = 0; // Store 0 if not positive/valid
                    }
                }
            }
             console.log(`Ítem #${itemNumber}: Materiales redondeados y sumados a total. Rounded quantities for this item:`, roundedItemMaterials);

        }); // End of itemBlocks.forEach

        console.log("Fin del procesamiento de ítems.");
        console.log("Errores totales encontrados:", currentErrorMessages);

        // --- Display Results ---
        if (currentErrorMessages.length > 0) {
            console.log("Mostrando mensajes de error.");
            // Display errors first if any
             resultsContent.innerHTML = '<div class="error-message"><h2>Errores de Validación:</h2>' +
                                        currentErrorMessages.map(msg => `<p>${msg}</p>`).join('') +
                                        '<p>Por favor, corrige los errores en los ítems marcados.</p></div>';
             // Clear/reset previous results and hide download buttons
             downloadOptionsDiv.classList.add('hidden');
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = currentErrorMessages; // Store errors for potential future handling
             return; // Stop here if there are errors
        }

        console.log("No se encontraron errores de validación.");
        // If no errors, proceed to display results and store them

        let resultsHtml = '<div class="report-header">';
        resultsHtml += '<h2>Resumen de Materiales</h2>';
        resultsHtml += `<p>Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}</p>`; // Format date for Spanish
        resultsHtml += '</div>';
        resultsHtml += '<hr>';

         // Display individual item summaries for valid items
        if (currentCalculatedItemsSpecs.length > 0) {
             console.log("Generando resumen de ítems calculados.");
             resultsHtml += '<h3>Detalle de Ítems Calculados:</h3>';
             currentCalculatedItemsSpecs.forEach(item => {
                 // Need the original item block element to check input visibility dynamically
                 const itemBlockElement = itemsContainer.querySelector(`[data-item-id="${item.id}"]`);
                 if (!itemBlockElement) return; // Should not happen if item is in currentCalculatedItemsSpecs, but safety check

                 const heightInputGroup = itemBlockElement.querySelector('.item-height-input');
                 const lengthInputGroup = itemBlockElement.querySelector('.item-length-input');
                 const facesInputGroup = itemBlockElement.querySelector('.item-faces-input');
                 const postSpacingInputGroup = itemBlockElement.querySelector('.item-post-spacing-input');
                 const plenumInputGroup = itemBlockElement.querySelector('.item-plenum-input');
                 const doubleStructureInputGroup = itemBlockElement.querySelector('.item-double-structure-input');


                 resultsHtml += `<div class="item-summary">`;
                 resultsHtml += `<h4>${getItemTypeName(item.type)} #${item.number}</h4>`;
                 resultsHtml += `<p><strong>Tipo:</strong> <span>${getItemTypeName(item.type)}</span></p>`;
                 resultsHtml += `<p><strong>Ancho:</strong> <span>${item.width.toFixed(2)} m</span></p>`;

                 if (item.type === 'muro-yeso' || item.type === 'muro-durock') {
                      if (!heightInputGroup.classList.contains('hidden')) resultsHtml += `<p><strong>Alto:</strong> <span>${item.height.toFixed(2)} m</span></p>`;
                      if (!facesInputGroup.classList.contains('hidden')) resultsHtml += `<p><strong>Nº Caras:</strong> <span>${item.faces}</span></p>`;
                      if (!postSpacingInputGroup.classList.contains('hidden')) resultsHtml += `<p><strong>Espaciamiento Postes:</strong> <span>${item.postSpacing.toFixed(2)} m</span></p>`;
                       if (!doubleStructureInputGroup.classList.contains('hidden')) resultsHtml += `<p><strong>Estructura Doble:</strong> <span>${item.isDoubleStructure ? 'Sí' : 'No'}</span></p>`;

                 } else if (item.type === 'cielo-yeso') {
                      if (!lengthInputGroup.classList.contains('hidden')) resultsHtml += `<p><strong>Largo:</strong> <span>${item.length.toFixed(2)} m</span></p>`;
                      if (!plenumInputGroup.classList.contains('hidden') && !isNaN(item.plenum)) resultsHtml += `<p><strong>Pleno:</strong> <span>${item.plenum.toFixed(2)} m</span></p>`;
                 }
                 resultsHtml += `</div>`;
             });
             resultsHtml += '<hr>';
        } else {
            console.log("No hay ítems calculados válidamente para mostrar resumen individual.");
        }


        resultsHtml += '<h3>Totales de Materiales (Cantidades a Comprar):</h3>';

        // currentTotalMaterials already holds summed rounded quantities.
        let hasMaterials = Object.keys(currentTotalMaterials).length > 0;
        console.log("Total de materiales calculados (sumas redondeadas):", currentTotalMaterials);

        if (hasMaterials) {
             console.log("Generando tabla de materiales totales.");
             // Sort materials alphabetically for consistent display
             const sortedMaterials = Object.keys(currentTotalMaterials).sort();

            resultsHtml += '<table><thead><tr><th>Material</th><th>Cantidad</th><th>Unidad</th></tr></thead><tbody>';

            sortedMaterials.forEach(material => {
                const cantidad = currentTotalMaterials[material];
                const unidad = getMaterialUnit(material); // Get unit using the helper function
                // Display the original material name
                 resultsHtml += `<tr><td>${material}</td><td>${cantidad}</td><td>${unidad}</td></tr>`;
            });

            resultsHtml += '</tbody></table>';
            downloadOptionsDiv.classList.remove('hidden'); // Show download options
        } else {
             console.log("No se calcularon materiales totales positivos.");
             // If there are no materials after calculation but no validation errors, something went wrong or inputs resulted in 0.
             resultsHtml += '<p>No se pudieron calcular los materiales con las dimensiones ingresadas. Revisa los valores.</p>';
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
        }

        // Append the generated HTML to the results content area
        resultsContent.innerHTML = resultsHtml; // Use = to replace previous content

        console.log("Resultados HTML generados y añadidos al DOM.");

        // Store the successfully calculated and rounded results and specs
        lastCalculatedTotalMaterials = currentTotalMaterials; // Store the final summed rounded values
        lastCalculatedItemsSpecs = currentCalculatedItemsSpecs; // Store specs of items included in calculation
        lastErrorMessages = []; // Clear errors as calculation was successful

         console.log("Estado de resultados almacenado para descarga.");
    };


    // --- PDF Generation Function ---
    const generatePDF = () => {
         console.log("Iniciando generación de PDF...");
        // Ensure there are calculated results to download
        if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
            console.warn("No hay resultados calculados para generar el PDF.");
            alert("Por favor, realiza un cálculo válido antes de generar el PDF.");
            return;
        }

        // Initialize jsPDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        // Define colors in RGB from CSS variables
        const primaryOliveRGB = [85, 107, 47]; // #556B2F
        const secondaryOliveRGB = [128, 128, 0]; // #808000
        const darkGrayRGB = [51, 51, 51]; // #333
        const mediumGrayRGB = [102, 102, 102]; // #666
        const lightGrayRGB = [224, 224, 224]; // #e0e0e0
        const extraLightGrayRGB = [248, 248, 248]; // #f8f8f8


        // --- Add Header ---
        doc.setFontSize(18);
        doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
        doc.setFont("helvetica", "bold"); // Use a standard font or include custom fonts
        doc.text("Resumen de Materiales Tablayeso", 14, 22);

        doc.setFontSize(10);
        doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
         doc.setFont("helvetica", "normal");
        doc.text(`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`, 14, 28);

        // Set starting Y position for the next content block
        let finalY = 35; // Start below the header

        // --- Add Item Summaries ---
        if (lastCalculatedItemsSpecs.length > 0) {
             console.log("Añadiendo resumen de ítems al PDF.");
             doc.setFontSize(14);
             doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
            doc.setFont("helvetica", "bold");
             doc.text("Detalle de Ítems Calculados:", 14, finalY + 10);
             finalY += 15; // Move Y below the title

             doc.setFontSize(9);
             doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
             doc.setFont("helvetica", "normal");

            const itemSummaryMargin = 5; // Space between summary lines

             lastCalculatedItemsSpecs.forEach(item => {
                 doc.setFontSize(10);
                 doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
                 doc.setFont("helvetica", "bold");
                 doc.text(`${getItemTypeName(item.type)} #${item.number}:`, 14, finalY + itemSummaryMargin);
                 finalY += itemSummaryMargin;

                 doc.setFontSize(9);
                 doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
                 doc.setFont("helvetica", "normal");

                 doc.text(`Tipo: ${getItemTypeName(item.type)}`, 20, finalY + itemSummaryMargin);
                 finalY += itemSummaryMargin;
                 doc.text(`Ancho: ${item.width.toFixed(2)} m`, 20, finalY + itemSummaryMargin);
                 finalY += itemSummaryMargin;

                 if (item.type === 'muro-yeso' || item.type === 'muro-durock') {
                      // Based on input visibility logic, height, faces, postSpacing, doubleStructure are always visible for walls
                      if (!isNaN(item.height)) {
                           doc.text(`Alto: ${item.height.toFixed(2)} m`, 20, finalY + itemSummaryMargin);
                           finalY += itemSummaryMargin;
                      }
                      if (!isNaN(item.faces)) {
                           doc.text(`Nº Caras: ${item.faces}`, 20, finalY + itemSummaryMargin);
                           finalY += itemSummaryMargin;
                      }
                       if (!isNaN(item.postSpacing)) {
                          doc.text(`Espaciamiento Postes: ${item.postSpacing.toFixed(2)} m`, 20, finalY + itemSummaryMargin);
                           finalY += itemSummaryMargin;
                      }
                       doc.text(`Estructura Doble: ${item.isDoubleStructure ? 'Sí' : 'No'}`, 20, finalY + itemSummaryMargin);
                        finalY += itemSummaryMargin;


                 } else if (item.type === 'cielo-yeso') {
                      // Based on input visibility logic, length, plenum are always visible for ceilings
                       if (!isNaN(item.length)) {
                           doc.text(`Largo: ${item.length.toFixed(2)} m`, 20, finalY + itemSummaryMargin);
                           finalY += itemSummaryMargin;
                       }
                     if (!isNaN(item.plenum)) { // Plenum needs valid number check
                         doc.text(`Pleno: ${item.plenum.toFixed(2)} m`, 20, finalY + itemSummaryMargin);
                          finalY += itemSummaryMargin;
                     }
                 }
                 finalY += itemSummaryMargin * 1.5; // Add a bit more space after each item summary
             });
             finalY += 5; // Add space before the total materials table title
        } else {
             console.log("No hay ítems calculados válidamente para añadir resumen al PDF.");
        }


        // --- Add Total Materials Table ---
         console.log("Añadiendo tabla de materiales totales al PDF.");
        doc.setFontSize(14);
        doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
        doc.setFont("helvetica", "bold");
        doc.text("Totales de Materiales:", 14, finalY + 10);
        finalY += 15; // Move Y below the title

        const tableColumn = ["Material", "Cantidad", "Unidad"];
        const tableRows = [];

        // Prepare data for the table
        const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
        sortedMaterials.forEach(material => {
            const cantidad = lastCalculatedTotalMaterials[material];
            const unidad = getMaterialUnit(material); // Get unit using the helper function
            // Use the material name directly from the key
             tableRows.push([material, cantidad, unidad]);
        });

         // Add the table using jspdf-autotable
         doc.autoTable({
             head: [tableColumn],
             body: tableRows,
             startY: finalY, // Start position below the last content
             theme: 'plain', // Start with a plain theme to apply custom styles
             headStyles: {
                 fillColor: lightGrayRGB,
                 textColor: darkGrayRGB,
                 fontStyle: 'bold',
                 halign: 'center', // Horizontal alignment
                 valign: 'middle', // Vertical alignment
                 lineWidth: 0.1,
                 lineColor: lightGrayRGB,
                 fontSize: 10 // Match HTML table header font size
             },
             bodyStyles: {
                 textColor: darkGrayRGB,
                 lineWidth: 0.1,
                 lineColor: lightGrayRGB,
                 fontSize: 9 // Match HTML table body font size
             },
              alternateRowStyles: { // Styling for alternate rows
                 fillColor: extraLightGrayRGB,
             },
              // Specific column styles (Cantidad column is the second one, index 1)
             columnStyles: {
                 1: {
                     halign: 'right', // Align quantity to the right
                     fontStyle: 'bold',
                     textColor: primaryOliveRGB // For quantity text color
                 },
                  2: { // Unit column
                     halign: 'center' // Align unit to the center or left as preferred
                 }
             },
             margin: { top: 10, right: 14, bottom: 14, left: 14 }, // Add margin
              didDrawPage: function (data) {
                // Optional: Add page number or footer here
                doc.setFontSize(8);
                doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
                doc.text('© 2025 [PROPUL] - Calculadora de Materiales Tablayeso v2.0', data.settings.margin.left, doc.internal.pageSize.height - 10);
             }
         });

         // Update finalY after the table
        finalY = doc.autoTable.previous.finalY;

        console.log("PDF generado.");
        // --- Save the PDF ---
        doc.save(`Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.pdf`); // Filename with date
    };


     // --- Excel Generation Function ---
    const generateExcel = () => {
         console.log("Iniciando generación de Excel...");
         // Ensure there are calculated results to download
        if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
            console.warn("No hay resultados calculados para generar el Excel.");
            alert("Por favor, realiza un cálculo válido antes de generar el Excel.");
            return;
        }

        // Data for the Excel sheet
        let sheetData = [];

        // Add Header
        sheetData.push(["Calculadora de Materiales Tablayeso"]);
        sheetData.push([`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`]);
        sheetData.push([]); // Blank row for spacing

        // Add Item Summaries (simplified for Excel)
         console.log("Añadiendo resumen de ítems al Excel.");
        sheetData.push(["Detalle de Ítems Calculados:"]);
        sheetData.push(["Tipo", "Número", "Ancho (m)", "Alto (m)", "Largo (m)", "Nº Caras", "Espaciamiento Postes (m)", "Pleno (m)", "Estructura Doble"]);
        lastCalculatedItemsSpecs.forEach(item => {
            // Based on input visibility logic, include data if type matches and it's a valid number
            sheetData.push([
                 getItemTypeName(item.type),
                 item.number,
                 item.width.toFixed(2),
                 (item.type === 'muro-yeso' || item.type === 'muro-durock') && !isNaN(item.height) ? item.height.toFixed(2) : '', // Only show if wall and valid number
                 item.type === 'cielo-yeso' && !isNaN(item.length) ? item.length.toFixed(2) : '', // Only show if ceiling and valid number
                 (item.type === 'muro-yeso' || item.type === 'muro-durock') && !isNaN(item.faces) ? item.faces : '', // Only show if wall and valid number
                 (item.type === 'muro-yeso' || item.type === 'muro-durock') && !isNaN(item.postSpacing) ? item.postSpacing.toFixed(2) : '', // Only show if wall and valid number
                 item.type === 'cielo-yeso' && !isNaN(item.plenum) ? item.plenum.toFixed(2) : '', // Only show if ceiling and valid number
                 (item.type === 'muro-yeso' || item.type === 'muro-durock') ? (item.isDoubleStructure ? 'Sí' : 'No') : '' // Only show if wall
             ]);
        });
         sheetData.push([]); // Blank row for spacing

        // Add Total Materials Table
         console.log("Añadiendo tabla de materiales totales al Excel.");
        sheetData.push(["Totales de Materiales (Cantidades a Comprar):"]);
        sheetData.push(["Material", "Cantidad", "Unidad"]);

        const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
        sortedMaterials.forEach(material => {
            const cantidad = lastCalculatedTotalMaterials[material];
            const unidad = getMaterialUnit(material);
            // Use the material name directly from the key
             sheetData.push([material, cantidad, unidad]);
        });

        // Create a workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(sheetData);

         // Auto-fit column widths - basic approach
         const colWidths = sheetData.reduce((acc, row) => {
             row.forEach((cell, i) => {
                 const cellValue = cell === null || cell === undefined ? '' : cell.toString();
                 const cellLength = cellValue.length;
                 acc[i] = Math.max(acc[i] || 0, cellLength);
             });
             return acc;
         }, []);
         ws['!cols'] = colWidths.map(w => ({ wch: w + 2 })); // Add a little extra padding


        // Append the worksheet to the workbook
        XLSX.utils.book_append_sheet(wb, ws, "Resumen Materiales");

         console.log("Excel generado.");
        // Write and download the Excel file
        XLSX.writeFile(wb, `Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.xlsx`); // Filename with date
    };


    // --- Event Listeners ---
    addItemBtn.addEventListener('click', createItemBlock);
    calculateBtn.addEventListener('click', calculateMaterials);
    generatePdfBtn.addEventListener('click', generatePDF);
    generateExcelBtn.addEventListener('click', generateExcel);


    // --- Initial State ---
    // Add one item block on page load
    createItemBlock();
     // Disable calculate button initially if no items are present (though createItemBlock adds one)
     toggleCalculateButtonState();
});
