// Variáveis globais
let currentFile = null;
let excelData = null;
let columnMapping = {};
let currentOperationType = null;

// Inicialização
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    loadSavedParams();
});

function initializeApp() {
    // Configurar drag and drop
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');

    uploadArea.addEventListener('click', () => fileInput.click());
    
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    });
    
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });
}

// Navegação entre páginas
function showPage(pageName) {
    // Esconder todas as páginas
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    
    // Remover classe active de todos os botões
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Mostrar página selecionada
    document.getElementById(pageName + '-page').classList.add('active');
    
    // Adicionar classe active ao botão correspondente
    event.target.classList.add('active');
}

// Manipulação de arquivos
function handleFileSelect(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert('Por favor, selecione um arquivo Excel (.xlsx ou .xls)');
        return;
    }
    
    currentFile = file;
    displayFileInfo(file);
    
    // Ler o arquivo Excel
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Converter para JSON
            excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (excelData.length > 0) {
                displayPreview(excelData);
                updateColumnSelects(excelData[0]);
                document.getElementById('convert-btn').disabled = false;
            } else {
                alert('O arquivo Excel está vazio ou não contém dados válidos.');
            }
        } catch (error) {
            console.error('Erro ao ler arquivo Excel:', error);
            alert('Erro ao ler o arquivo Excel. Verifique se o arquivo não está corrompido.');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function displayFileInfo(file) {
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'block';
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function removeFile() {
    currentFile = null;
    excelData = null;
    
    document.getElementById('file-info').style.display = 'none';
    document.getElementById('preview-container').style.display = 'none';
    document.getElementById('convert-btn').disabled = true;
    document.getElementById('file-input').value = '';
}

function displayPreview(data) {
    const previewContainer = document.getElementById('preview-container');
    const previewHeader = document.getElementById('preview-header');
    const previewBody = document.getElementById('preview-body');
    
    // Limpar conteúdo anterior
    previewHeader.innerHTML = '';
    previewBody.innerHTML = '';
    
    if (data.length === 0) return;
    
    // Criar cabeçalho
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header || 'Coluna ' + (headerRow.children.length + 1);
        headerRow.appendChild(th);
    });
    previewHeader.appendChild(headerRow);
    
    // Criar linhas de dados (máximo 10 linhas para preview)
    const maxRows = Math.min(data.length - 1, 10);
    for (let i = 1; i <= maxRows; i++) {
        const row = document.createElement('tr');
        data[i].forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            row.appendChild(td);
        });
        previewBody.appendChild(row);
    }
    
    previewContainer.style.display = 'block';
}

// Gerenciamento de parâmetros
function updateColumnSelects(headers) {
    const selects = document.querySelectorAll('.param-select');
    
    selects.forEach(select => {
        // Manter a opção padrão
        const defaultOption = select.querySelector('option[value=""]');
        select.innerHTML = '';
        select.appendChild(defaultOption);
        
        // Adicionar opções das colunas
        headers.forEach((header, index) => {
            if (header) {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = header;
                select.appendChild(option);
            }
        });
    });
}

// Função para mostrar parâmetros baseado no tipo de operação
function showOperationParams() {
    const operationType = document.querySelector('input[name="operation-type"]:checked');
    
    if (!operationType) {
        // Esconder todos os parâmetros se nenhum tipo for selecionado
        document.querySelectorAll('.operation-params').forEach(params => {
            params.style.display = 'none';
        });
        return;
    }
    
    currentOperationType = operationType.value;
    
    // Esconder todos os parâmetros primeiro
    document.querySelectorAll('.operation-params').forEach(params => {
        params.style.display = 'none';
    });
    
    // Mostrar apenas os parâmetros do tipo selecionado
    const targetParams = document.getElementById(operationType.value + '-params');
    if (targetParams) {
        targetParams.style.display = 'block';
    }
    
    // Atualizar os selects se houver dados carregados
    if (excelData && excelData.length > 0) {
        updateColumnSelects(excelData[0]);
    }
}

function saveParamsWithName() {
    if (!currentOperationType) {
        showNotification('Por favor, selecione um tipo de operação primeiro!', 'warning');
        return;
    }
    
    const params = {};
    const activeParams = document.querySelector('.operation-params[style*="block"]');
    
    if (activeParams) {
        const selects = activeParams.querySelectorAll('.param-select');
        
        selects.forEach(select => {
            const fieldName = select.id.replace('-col', '');
            params[fieldName] = select.value;
        });
    }
    
    // Solicitar nome para os parâmetros
    const paramName = prompt('Digite um nome para salvar estes parâmetros:');
    
    if (!paramName || paramName.trim() === '') {
        showNotification('Nome é obrigatório para salvar os parâmetros!', 'warning');
        return;
    }
    
    // Salvar no localStorage com nome e tipo de operação
    const savedData = {
        name: paramName.trim(),
        operationType: currentOperationType,
        params: params,
        date: new Date().toISOString()
    };
    
    // Obter lista existente de parâmetros salvos
    const savedParamsList = JSON.parse(localStorage.getItem('excelConverterParamsList') || '[]');
    
    // Verificar se já existe um parâmetro com este nome
    const existingIndex = savedParamsList.findIndex(item => item.name === paramName.trim());
    
    if (existingIndex !== -1) {
        if (confirm('Já existe um parâmetro com este nome. Deseja substituí-lo?')) {
            savedParamsList[existingIndex] = savedData;
        } else {
            return;
        }
    } else {
        savedParamsList.push(savedData);
    }
    
    // Salvar lista atualizada
    localStorage.setItem('excelConverterParamsList', JSON.stringify(savedParamsList));
    
    // Salvar como parâmetros atuais também
    localStorage.setItem('excelConverterParams', JSON.stringify(savedData));
    columnMapping = params;
    
    showNotification(`Parâmetros "${paramName}" salvos com sucesso!`, 'success');
}

function loadSavedParamsList() {
    const savedParamsList = JSON.parse(localStorage.getItem('excelConverterParamsList') || '[]');
    const savedParamsSection = document.getElementById('saved-params-section');
    const savedParamsListElement = document.getElementById('saved-params-list');
    
    if (savedParamsList.length === 0) {
        showNotification('Nenhum parâmetro salvo encontrado!', 'info');
        return;
    }
    
    // Limpar lista atual
    savedParamsListElement.innerHTML = '';
    
    // Adicionar cada parâmetro salvo
    savedParamsList.forEach((savedParam, index) => {
        const paramItem = document.createElement('div');
        paramItem.className = 'saved-param-item';
        
        const operationTypeLabels = {
            'titular': 'Titular',
            'dependente': 'Dependente',
            'geral': 'Geral'
        };
        
        paramItem.innerHTML = `
            <div class="saved-param-header">
                <span class="saved-param-name">${savedParam.name}</span>
                <span class="saved-param-type">${operationTypeLabels[savedParam.operationType] || savedParam.operationType}</span>
            </div>
            <div class="saved-param-date">Salvo em: ${new Date(savedParam.date).toLocaleDateString('pt-BR')}</div>
            <div class="saved-param-actions">
                <button class="saved-param-btn load" onclick="loadSpecificParams(${index})">
                    <i class="fas fa-download"></i> Carregar
                </button>
                <button class="saved-param-btn delete" onclick="deleteSavedParams(${index})">
                    <i class="fas fa-trash"></i> Excluir
                </button>
            </div>
        `;
        
        savedParamsListElement.appendChild(paramItem);
    });
    
    // Mostrar seção
    savedParamsSection.style.display = 'block';
}

function loadSpecificParams(index) {
    const savedParamsList = JSON.parse(localStorage.getItem('excelConverterParamsList') || '[]');
    const savedParam = savedParamsList[index];
    
    if (!savedParam) {
        showNotification('Parâmetro não encontrado!', 'error');
        return;
    }
    
    // Selecionar tipo de operação
    const radioButton = document.querySelector(`input[name="operation-type"][value="${savedParam.operationType}"]`);
    if (radioButton) {
        radioButton.checked = true;
        showOperationParams();
    }
    
    // Aplicar parâmetros após um pequeno delay para garantir que os selects estejam carregados
    setTimeout(() => {
        Object.keys(savedParam.params).forEach(fieldName => {
            const select = document.getElementById(fieldName + '-col');
            if (select) {
                select.value = savedParam.params[fieldName];
            }
        });
        
        columnMapping = savedParam.params;
        currentOperationType = savedParam.operationType;
        
        showNotification(`Parâmetros "${savedParam.name}" carregados com sucesso!`, 'success');
    }, 100);
}

function deleteSavedParams(index) {
    const savedParamsList = JSON.parse(localStorage.getItem('excelConverterParamsList') || '[]');
    const savedParam = savedParamsList[index];
    
    if (confirm(`Tem certeza que deseja excluir os parâmetros "${savedParam.name}"?`)) {
        savedParamsList.splice(index, 1);
        localStorage.setItem('excelConverterParamsList', JSON.stringify(savedParamsList));
        
        showNotification(`Parâmetros "${savedParam.name}" excluídos com sucesso!`, 'success');
        
        // Recarregar lista
        loadSavedParamsList();
    }
}

function loadSavedParams() {
    const saved = localStorage.getItem('excelConverterParams');
    if (saved) {
        const savedData = JSON.parse(saved);
        
        // Se for formato antigo (sem operationType), converter
        if (savedData.operationType) {
            currentOperationType = savedData.operationType;
            columnMapping = savedData.params;
            
            // Selecionar o tipo de operação
            const radioButton = document.querySelector(`input[name="operation-type"][value="${currentOperationType}"]`);
            if (radioButton) {
                radioButton.checked = true;
                showOperationParams();
            }
            
            // Aplicar valores salvos aos selects
            setTimeout(() => {
                Object.keys(columnMapping).forEach(fieldName => {
                    const select = document.getElementById(fieldName + '-col');
                    if (select) {
                        select.value = columnMapping[fieldName];
                    }
                });
            }, 100);
        } else {
            // Formato antigo - converter para novo formato
            columnMapping = savedData;
            currentOperationType = 'geral';
            
            const radioButton = document.querySelector('input[name="operation-type"][value="geral"]');
            if (radioButton) {
                radioButton.checked = true;
                showOperationParams();
            }
        }
    }
}

function resetParams() {
    if (confirm('Tem certeza que deseja resetar todos os parâmetros?')) {
        localStorage.removeItem('excelConverterParams');
        columnMapping = {};
        currentOperationType = null;
        
        // Limpar seleção de tipo de operação
        document.querySelectorAll('input[name="operation-type"]').forEach(radio => {
            radio.checked = false;
        });
        
        // Esconder todos os parâmetros
        document.querySelectorAll('.operation-params').forEach(params => {
            params.style.display = 'none';
        });
        
        // Limpar todos os selects
        const selects = document.querySelectorAll('.param-select');
        selects.forEach(select => {
            select.value = '';
        });
        
        showNotification('Parâmetros resetados!', 'info');
    }
}

// Conversão de arquivo
function convertFile() {
    if (!currentFile || !excelData) {
        alert('Por favor, selecione um arquivo Excel primeiro.');
        return;
    }
    
    if (!currentOperationType) {
        alert('Por favor, selecione um tipo de operação primeiro.');
        return;
    }
    
    if (Object.keys(columnMapping).length === 0) {
        alert('Por favor, configure os parâmetros de mapeamento primeiro.');
        return;
    }
    
    // Verificar se os parâmetros obrigatórios estão configurados
    const requiredFields = getRequiredFieldsForOperation(currentOperationType);
    const missingFields = requiredFields.filter(field => !columnMapping[field] || columnMapping[field] === '');
    
    if (missingFields.length > 0) {
        alert(`Por favor, configure os seguintes campos obrigatórios: ${missingFields.join(', ')}`);
        return;
    }
    
    showProgressModal();
    
    // Simular processamento
    setTimeout(() => {
        const convertedData = processExcelData(excelData, columnMapping, currentOperationType);
        
        // Se for operação de titular, mostrar informações sobre operadoras e CPFs inválidos
        if (currentOperationType === 'titular' && typeof convertedData === 'object' && !Array.isArray(convertedData)) {
            const operadoras = Object.keys(convertedData);
            const operadoraInfo = operadoras.map(op => {
                const count = convertedData[op].length - 1; // -1 para excluir cabeçalho
                return `${op} (${count} registros)`;
            }).join(', ');
            
            // Contar CPFs inválidos e dependentes órfãos totais
            let totalCpfsInvalidos = 0;
            let totalDependentesOrfaos = 0;
            operadoras.forEach(operadora => {
                totalCpfsInvalidos += countInvalidCpfs(convertedData[operadora]);
                totalDependentesOrfaos += countOrphanDependents(convertedData[operadora]);
            });
            
            let message = `Detectadas ${operadoras.length} operadora(s): ${operadoraInfo}`;
            let hasIssues = false;
            
            if (totalCpfsInvalidos > 0) {
                message += `. ⚠️ ${totalCpfsInvalidos} CPF(s) inválido(s)`;
                hasIssues = true;
            }
            
            if (totalDependentesOrfaos > 0) {
                message += `. ⚠️ ${totalDependentesOrfaos} dependente(s) sem titular`;
                hasIssues = true;
            }
            
            showNotification(message, hasIssues ? 'warning' : 'info');
        }
        
        downloadConvertedFile(convertedData, currentFile.name);
        hideProgressModal();
    }, 2000);
}

function getRequiredFieldsForOperation(operationType) {
    switch (operationType) {
        case 'titular':
            return ['operadora', 'nome-segurado', 'cpf', 'data-nascimento'];
        case 'dependente':
            return ['operadora', 'nome-dependente', 'cpf-dependente', 'cpf-titular', 'nome-titular'];
        case 'exclusao':
            return ['operadora', 'tipo-exclusao', 'matricula'];
        case 'geral':
            return ['nome'];
        default:
            return [];
    }
}

function processExcelData(data, mapping, operationType) {
    const headers = data[0];
    
    // Definir cabeçalhos baseado no tipo de operação
    let newHeaders = [];
    switch (operationType) {
        case 'titular':
            newHeaders = [
                'Operadora', 'Nº da apólice/contrato', 'Nº Sub', 'Inicio de Vigência', 'Início de Vigência Programada',
                'Código/Nome do Plano', 'Matrícula', 'Código Dependente', 'Data de admissão', 'Nome Segurado',
                'Sexo', 'Data de Nascimento', 'Nº CPF', 'Estado Civil', 'Grau de Parentesco', 'Peso', 'Altura',
                'Data da Tutela', 'Data do Documento', 'Data do Casamento', 'Nome da Mãe', 'Nome do Pai', 'PIS',
                'RG', 'Órgão Emissor', 'Data Expedição', 'Nº Centro de Custo', 'Nome Centro de Custo', 'Cargo',
                'Tipo de Logradouro', 'Nome do Logradouro', 'Número', 'Complemento', 'Bairro', 'Cidade', 'Estado',
                'CEP', 'Cartão nacional de saúde', 'Banco', 'Agência', 'Dígito da Agência', 'Conta Corrente',
                'Dígito da conta corrente', 'Número de nascido vivo', 'Telefone', 'E-mail', 'Lote ou Chamado',
                'Documento', 'Sequência', 'Setor', 'Lotação', 'Local', 'Unidade de Atendimento', 'Unidade de Negociação'
            ];
            break;
        case 'dependente':
            newHeaders = [
                'Operadora', 'Nº da apólice/contrato', 'Nº Sub', 'Matrícula (Titular)', 'Certificado',
                'Código Dependente', 'Nome do Titular', 'CPF Titular', 'Nome Dependente', 'CPF do Dependente',
                'Data de Nascimento', 'Sexo', 'Estado Civil', 'Grau de Parentesco', 'Peso', 'Altura',
                'Data da Tutela', 'Data do Documento', 'Data do Casamento', 'Inicio de Vigência',
                'Início de Vigência Programada', 'Nome da Mãe do dependente', 'Nome do Pai',
                'Número de nascido vivo', 'Cartão nacional de saúde (CNS)', 'Lote ou Chamado',
                'Documento', 'Sequencia', 'Setor', 'Lotação', 'Local'
            ];
            break;
        case 'exclusao':
            newHeaders = [
                'Tipo de Exclusão', 'Nº da apólice/contrato', 'Nº Sub', 'Operadora', 'Matrícula',
                'Certificado', 'Código Dependente', 'Nome do Titular', 'CPF Titular', 'Data de Demissão',
                'Nome Dependente', 'Data de Cancelamento', 'Data de Cancelamento Programada', 'Motivo',
                'Lote ou Chamado', 'Documento', 'Sequencia', 'Setor', 'Lotação', 'Local'
            ];
            break;
        case 'geral':
            newHeaders = Object.keys(mapping).map(field => {
                const columnIndex = mapping[field];
                return columnIndex !== '' ? headers[columnIndex] : field;
            });
            break;
        default:
            newHeaders = Object.keys(mapping).map(field => {
                const columnIndex = mapping[field];
                return columnIndex !== '' ? headers[columnIndex] : field;
            });
    }
    
    // Se for operação de titular, separar por operadora
    if (operationType === 'titular' && mapping['operadora']) {
        return processDataByOperadora(data, mapping, newHeaders);
    }
    
    // Para outros tipos, processar normalmente
    const convertedData = [newHeaders];
    
    // Processar linhas de dados
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let newRow = [];
        
        switch (operationType) {
            case 'titular':
                // Obter CPF e validar
                const cpf = row[mapping['cpf']] || '';
                const cpfValido = isValidCPF(cpf);
                const cpfFormatado = cpfValido ? cpf : (cpf ? `[INVÁLIDO] ${cpf}` : '');
                
                // Verificar se é dependente órfão
                const matricula = row[mapping['matricula']] || '';
                const grauParentesco = row[mapping['grau-parentesco']] || '';
                const isDependente = grauParentesco && !grauParentesco.toLowerCase().includes('titular');
                const isOrphan = isDependente && !hasTitularInData(excelData, matricula, mapping);
                
                // Formatar matrícula se for dependente órfão
                const matriculaFormatada = isOrphan ? `[SEM TITULAR] ${matricula}` : matricula;
                
                newRow = [
                    row[mapping['operadora']] || '',
                    row[mapping['apolice']] || '',
                    row[mapping['sub']] || '',
                    formatDate(row[mapping['inicio-vigencia']]),
                    formatDate(row[mapping['inicio-vigencia-programada']]),
                    row[mapping['codigo-plano']] || '',
                    matriculaFormatada,
                    row[mapping['codigo-dependente']] || '',
                    formatDate(row[mapping['data-admissao']]),
                    row[mapping['nome-segurado']] || '',
                    row[mapping['sexo']] || '',
                    formatDate(row[mapping['data-nascimento']]),
                    cpfFormatado,
                    row[mapping['estado-civil']] || '',
                    grauParentesco,
                    row[mapping['peso']] || '',
                    row[mapping['altura']] || '',
                    formatDate(row[mapping['data-tutela']]),
                    formatDate(row[mapping['data-documento']]),
                    formatDate(row[mapping['data-casamento']]),
                    row[mapping['nome-mae']] || '',
                    row[mapping['nome-pai']] || '',
                    row[mapping['pis']] || '',
                    row[mapping['rg']] || '',
                    row[mapping['orgao-emissor']] || '',
                    formatDate(row[mapping['data-expedicao']]),
                    row[mapping['centro-custo']] || '',
                    row[mapping['nome-centro-custo']] || '',
                    row[mapping['cargo']] || '',
                    row[mapping['tipo-logradouro']] || '',
                    row[mapping['nome-logradouro']] || '',
                    row[mapping['numero']] || '',
                    row[mapping['complemento']] || '',
                    row[mapping['bairro']] || '',
                    row[mapping['cidade']] || '',
                    row[mapping['estado']] || '',
                    row[mapping['cep']] || '',
                    row[mapping['cartao-saude']] || '',
                    row[mapping['banco']] || '',
                    row[mapping['agencia']] || '',
                    row[mapping['digito-agencia']] || '',
                    row[mapping['conta-corrente']] || '',
                    row[mapping['digito-conta']] || '',
                    row[mapping['numero-nascido-vivo']] || '',
                    row[mapping['telefone']] || '',
                    row[mapping['email']] || '',
                    row[mapping['lote-chamado']] || '',
                    row[mapping['documento']] || '',
                    row[mapping['sequencia']] || '',
                    row[mapping['setor']] || '',
                    row[mapping['lotacao']] || '',
                    row[mapping['local']] || '',
                    row[mapping['unidade-atendimento']] || '',
                    row[mapping['unidade-negociacao']] || ''
                ];
                break;
            case 'dependente':
                // Obter CPF do dependente e validar
                const cpfDependente = row[mapping['cpf-dependente']] || '';
                const cpfDependenteValido = isValidCPF(cpfDependente);
                const cpfDependenteFormatado = cpfDependenteValido ? cpfDependente : (cpfDependente ? `[INVÁLIDO] ${cpfDependente}` : '');
                
                // Obter CPF do titular e validar
                const cpfTitular = row[mapping['cpf-titular']] || '';
                const cpfTitularValido = isValidCPF(cpfTitular);
                const cpfTitularFormatado = cpfTitularValido ? cpfTitular : (cpfTitular ? `[INVÁLIDO] ${cpfTitular}` : '');
                
                newRow = [
                    row[mapping['operadora']] || '',
                    row[mapping['apolice']] || '',
                    row[mapping['sub']] || '',
                    row[mapping['matricula-titular']] || '',
                    row[mapping['certificado']] || '',
                    row[mapping['codigo-dependente']] || '',
                    row[mapping['nome-titular']] || '',
                    cpfTitularFormatado,
                    row[mapping['nome-dependente']] || '',
                    cpfDependenteFormatado,
                    formatDate(row[mapping['data-nascimento-dependente']]),
                    row[mapping['sexo-dependente']] || '',
                    row[mapping['estado-civil-dependente']] || '',
                    row[mapping['grau-parentesco']] || '',
                    row[mapping['peso-dependente']] || '',
                    row[mapping['altura-dependente']] || '',
                    formatDate(row[mapping['data-tutela-dependente']]),
                    formatDate(row[mapping['data-documento-dependente']]),
                    formatDate(row[mapping['data-casamento-dependente']]),
                    formatDate(row[mapping['inicio-vigencia-dependente']]),
                    formatDate(row[mapping['inicio-vigencia-programada-dependente']]),
                    row[mapping['nome-mae-dependente']] || '',
                    row[mapping['nome-pai-dependente']] || '',
                    row[mapping['numero-nascido-vivo-dependente']] || '',
                    row[mapping['cartao-saude-dependente']] || '',
                    row[mapping['lote-chamado-dependente']] || '',
                    row[mapping['documento-dependente']] || '',
                    row[mapping['sequencia-dependente']] || '',
                    row[mapping['setor-dependente']] || '',
                    row[mapping['lotacao-dependente']] || '',
                    row[mapping['local-dependente']] || ''
                ];
                break;
            case 'exclusao':
                // Obter CPF do titular e validar
                const cpfTitularExclusao = row[mapping['cpf-titular-exclusao']] || '';
                const cpfTitularExclusaoValido = isValidCPF(cpfTitularExclusao);
                const cpfTitularExclusaoFormatado = cpfTitularExclusaoValido ? cpfTitularExclusao : (cpfTitularExclusao ? `[INVÁLIDO] ${cpfTitularExclusao}` : '');
                
                newRow = [
                    row[mapping['tipo-exclusao']] || '',
                    row[mapping['apolice-exclusao']] || '',
                    row[mapping['sub-exclusao']] || '',
                    row[mapping['operadora-exclusao']] || '',
                    row[mapping['matricula-exclusao']] || '',
                    row[mapping['certificado-exclusao']] || '',
                    row[mapping['codigo-dependente-exclusao']] || '',
                    row[mapping['nome-titular-exclusao']] || '',
                    cpfTitularExclusaoFormatado,
                    formatDate(row[mapping['data-demissao']]),
                    row[mapping['nome-dependente-exclusao']] || '',
                    formatDate(row[mapping['data-cancelamento']]),
                    formatDate(row[mapping['data-cancelamento-programada']]),
                    row[mapping['motivo']] || '',
                    row[mapping['lote-chamado-exclusao']] || '',
                    row[mapping['documento-exclusao']] || '',
                    row[mapping['sequencia-exclusao']] || '',
                    row[mapping['setor-exclusao']] || '',
                    row[mapping['lotacao-exclusao']] || '',
                    row[mapping['local-exclusao']] || ''
                ];
                break;
            case 'geral':
                newRow = Object.keys(mapping).map(field => {
                    const columnIndex = mapping[field];
                    const value = columnIndex !== '' ? (row[columnIndex] || '') : '';
                    // Verificar se é um campo de data baseado no nome do campo ou detecção automática
                    return (isDateField(field) || detectDateColumn(excelData, columnIndex)) ? formatDate(value) : value;
                });
                break;
            default:
                newRow = Object.keys(mapping).map(field => {
                    const columnIndex = mapping[field];
                    const value = columnIndex !== '' ? (row[columnIndex] || '') : '';
                    // Verificar se é um campo de data baseado no nome do campo ou detecção automática
                    return (isDateField(field) || detectDateColumn(excelData, columnIndex)) ? formatDate(value) : value;
                });
        }
        
        convertedData.push(newRow);
    }
    
    return convertedData;
}

// Função para processar dados separados por operadora
function processDataByOperadora(data, mapping, headers) {
    const operadoras = {};
    
    // Processar linhas de dados e agrupar por operadora
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const operadora = row[mapping['operadora']] || 'Sem Operadora';
        
        // Obter CPF e validar
        const cpf = row[mapping['cpf']] || '';
        const cpfValido = isValidCPF(cpf);
        const cpfFormatado = cpfValido ? cpf : (cpf ? `[INVÁLIDO] ${cpf}` : '');
        
        // Verificar se é dependente órfão
        const matricula = row[mapping['matricula']] || '';
        const grauParentesco = row[mapping['grau-parentesco']] || '';
        const isDependente = grauParentesco && !grauParentesco.toLowerCase().includes('titular');
        const isOrphan = isDependente && !hasTitularInData(data, matricula, mapping);
        
        // Formatar matrícula se for dependente órfão
        const matriculaFormatada = isOrphan ? `[SEM TITULAR] ${matricula}` : matricula;
        
        // Criar linha processada
        const newRow = [
            row[mapping['operadora']] || '',
            row[mapping['apolice']] || '',
            row[mapping['sub']] || '',
            formatDate(row[mapping['inicio-vigencia']]),
            formatDate(row[mapping['inicio-vigencia-programada']]),
            row[mapping['codigo-plano']] || '',
            matriculaFormatada,
            row[mapping['codigo-dependente']] || '',
            formatDate(row[mapping['data-admissao']]),
            row[mapping['nome-segurado']] || '',
            row[mapping['sexo']] || '',
            formatDate(row[mapping['data-nascimento']]),
            cpfFormatado,
            row[mapping['estado-civil']] || '',
            grauParentesco,
            row[mapping['peso']] || '',
            row[mapping['altura']] || '',
            formatDate(row[mapping['data-tutela']]),
            formatDate(row[mapping['data-documento']]),
            formatDate(row[mapping['data-casamento']]),
            row[mapping['nome-mae']] || '',
            row[mapping['nome-pai']] || '',
            row[mapping['pis']] || '',
            row[mapping['rg']] || '',
            row[mapping['orgao-emissor']] || '',
            formatDate(row[mapping['data-expedicao']]),
            row[mapping['centro-custo']] || '',
            row[mapping['nome-centro-custo']] || '',
            row[mapping['cargo']] || '',
            row[mapping['tipo-logradouro']] || '',
            row[mapping['nome-logradouro']] || '',
            row[mapping['numero']] || '',
            row[mapping['complemento']] || '',
            row[mapping['bairro']] || '',
            row[mapping['cidade']] || '',
            row[mapping['estado']] || '',
            row[mapping['cep']] || '',
            row[mapping['cartao-saude']] || '',
            row[mapping['banco']] || '',
            row[mapping['agencia']] || '',
            row[mapping['digito-agencia']] || '',
            row[mapping['conta-corrente']] || '',
            row[mapping['digito-conta']] || '',
            row[mapping['numero-nascido-vivo']] || '',
            row[mapping['telefone']] || '',
            row[mapping['email']] || '',
            row[mapping['lote-chamado']] || '',
            row[mapping['documento']] || '',
            row[mapping['sequencia']] || '',
            row[mapping['setor']] || '',
            row[mapping['lotacao']] || '',
            row[mapping['local']] || '',
            row[mapping['unidade-atendimento']] || '',
            row[mapping['unidade-negociacao']] || ''
        ];
        
        // Adicionar à operadora correspondente
        if (!operadoras[operadora]) {
            operadoras[operadora] = [headers];
        }
        operadoras[operadora].push(newRow);
    }
    
    return operadoras;
}

// Função para formatar datas para dd/mm/aaaa
function formatDate(dateValue) {
    if (!dateValue || dateValue === '') return '';
    
    let date;
    
    // Se já é uma string no formato dd/mm/aaaa, retornar como está
    if (typeof dateValue === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(dateValue)) {
        return dateValue;
    }
    
    // Se já é uma string no formato dd/mm/aa, converter para dd/mm/aaaa
    if (typeof dateValue === 'string' && /^\d{2}\/\d{2}\/\d{2}$/.test(dateValue)) {
        const parts = dateValue.split('/');
        const year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
        return `${parts[0]}/${parts[1]}/${year}`;
    }
    
    // Se é um número (Excel date serial), converter
    if (typeof dateValue === 'number') {
        // Excel usa 1 de janeiro de 1900 como data base
        const excelEpoch = new Date(1900, 0, 1);
        date = new Date(excelEpoch.getTime() + (dateValue - 1) * 24 * 60 * 60 * 1000);
    } else {
        // Tentar criar uma data a partir do valor
        date = new Date(dateValue);
    }
    
    // Verificar se a data é válida
    if (isNaN(date.getTime())) {
        return dateValue; // Retornar o valor original se não conseguir converter
    }
    
    // Formatar para dd/mm/aaaa
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
}

// Função para identificar campos de data baseado no nome
function isDateField(fieldName) {
    const dateKeywords = [
        'data', 'date', 'nascimento', 'admissao', 'vigencia', 'tutela', 
        'documento', 'casamento', 'expedicao', 'inicio', 'fim'
    ];
    
    const lowerFieldName = fieldName.toLowerCase();
    return dateKeywords.some(keyword => lowerFieldName.includes(keyword));
}

// Função para detectar automaticamente se uma coluna contém datas
function detectDateColumn(data, columnIndex) {
    if (!data || data.length < 2) return false;
    
    let dateCount = 0;
    let totalCount = 0;
    
    // Verificar as primeiras 10 linhas (ou todas se menos que 10)
    const rowsToCheck = Math.min(data.length - 1, 10);
    
    for (let i = 1; i <= rowsToCheck; i++) {
        const value = data[i][columnIndex];
        if (value && value !== '') {
            totalCount++;
            
            // Verificar se é uma data válida
            if (isValidDate(value)) {
                dateCount++;
            }
        }
    }
    
    // Se pelo menos 70% dos valores são datas, considerar como coluna de data
    return totalCount > 0 && (dateCount / totalCount) >= 0.7;
}

// Função para verificar se um valor é uma data válida
function isValidDate(value) {
    if (!value) return false;
    
    // Se é um número (Excel date serial)
    if (typeof value === 'number') {
        return value > 1 && value < 100000; // Range válido para datas Excel
    }
    
    // Se é uma string, verificar formatos comuns
    if (typeof value === 'string') {
        // Formato dd/mm/aaaa ou dd/mm/aa
        if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(value)) {
            return true;
        }
        
        // Formato aaaa-mm-dd
        if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(value)) {
            return true;
        }
        
        // Formato dd-mm-aaaa
        if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(value)) {
            return true;
        }
    }
    
    // Tentar criar uma data
    const date = new Date(value);
    return !isNaN(date.getTime());
}

function downloadConvertedFile(data, originalFileName) {
    // Criar workbook
    const workbook = XLSX.utils.book_new();
    
    // Verificar se os dados são separados por operadora (objeto) ou dados normais (array)
    if (typeof data === 'object' && !Array.isArray(data)) {
        // Dados separados por operadora
        const operadoras = Object.keys(data);
        
        if (operadoras.length === 1) {
            // Se há apenas uma operadora, criar uma única aba
            const worksheet = XLSX.utils.aoa_to_sheet(data[operadoras[0]]);
            const sheetName = sanitizeSheetName(operadoras[0]);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        } else {
            // Se há múltiplas operadoras, criar abas separadas
            operadoras.forEach(operadora => {
                const worksheet = XLSX.utils.aoa_to_sheet(data[operadora]);
                const sheetName = sanitizeSheetName(operadora);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            });
            
            // Criar aba de resumo
            const resumoData = createResumoSheet(operadoras, data);
            const resumoWorksheet = XLSX.utils.aoa_to_sheet(resumoData);
            XLSX.utils.book_append_sheet(workbook, resumoWorksheet, 'Resumo');
        }
        
        showNotification(`Arquivo convertido com ${operadoras.length} operadora(s) e baixado com sucesso!`, 'success');
    } else {
        // Dados normais (array)
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Dados Convertidos');
        showNotification('Arquivo convertido e baixado com sucesso!', 'success');
    }
    
    // Gerar nome do arquivo
    const fileName = originalFileName.replace(/\.(xlsx|xls)$/, '_convertido.xlsx');
    
    // Download
    XLSX.writeFile(workbook, fileName);
}

// Função para sanitizar nomes de abas (remover caracteres inválidos)
function sanitizeSheetName(name) {
    // Remover caracteres inválidos para nomes de abas do Excel
    let sanitizedName = name.replace(/[\[\]\*\/\\\?\:]/g, '');
    
    // Limitar a 31 caracteres (limite do Excel)
    if (sanitizedName.length > 31) {
        sanitizedName = sanitizedName.substring(0, 31);
    }
    
    // Se ficou vazio, usar nome padrão
    if (!sanitizedName.trim()) {
        sanitizedName = 'Operadora';
    }
    
    return sanitizedName;
}

// Função para criar aba de resumo
function createResumoSheet(operadoras, data) {
    const resumoData = [
        ['Resumo por Operadora'],
        [''],
        ['Operadora', 'Quantidade de Registros', 'CPFs Inválidos', 'Dependentes Órfãos']
    ];
    
    let totalRegistros = 0;
    let totalCpfsInvalidos = 0;
    let totalDependentesOrfaos = 0;
    
    operadoras.forEach(operadora => {
        const quantidade = data[operadora].length - 1; // -1 para excluir o cabeçalho
        const cpfsInvalidos = countInvalidCpfs(data[operadora]);
        const dependentesOrfaos = countOrphanDependents(data[operadora]);
        
        resumoData.push([operadora, quantidade, cpfsInvalidos, dependentesOrfaos]);
        
        totalRegistros += quantidade;
        totalCpfsInvalidos += cpfsInvalidos;
        totalDependentesOrfaos += dependentesOrfaos;
    });
    
    // Adicionar total
    resumoData.push(['', '']);
    resumoData.push(['TOTAL', totalRegistros, totalCpfsInvalidos, totalDependentesOrfaos]);
    
    // Adicionar informações sobre problemas encontrados
    let hasIssues = false;
    let issuesText = [];
    
    if (totalCpfsInvalidos > 0) {
        hasIssues = true;
        issuesText.push('CPFs inválidos');
    }
    
    if (totalDependentesOrfaos > 0) {
        hasIssues = true;
        issuesText.push('dependentes sem titular');
    }
    
    if (hasIssues) {
        resumoData.push(['', '']);
        resumoData.push([`ATENÇÃO: Foram encontrados ${issuesText.join(' e ')} no arquivo!`]);
        resumoData.push(['Verifique os dados antes de prosseguir com a importação.']);
    }
    
    return resumoData;
}

// Função para contar CPFs inválidos em uma operadora
function countInvalidCpfs(operadoraData) {
    if (!operadoraData || operadoraData.length < 2) return 0;
    
    const cpfColumnIndex = 12; // Índice da coluna CPF (baseado na estrutura do titular)
    let invalidCount = 0;
    
    for (let i = 1; i < operadoraData.length; i++) {
        const cpf = operadoraData[i][cpfColumnIndex];
        if (cpf && !isValidCPF(cpf)) {
            invalidCount++;
        }
    }
    
    return invalidCount;
}

// Função para validar CPF
function isValidCPF(cpf) {
    if (!cpf) return false;
    
    // Remover caracteres não numéricos
    cpf = cpf.replace(/[^\d]/g, '');
    
    // Verificar se tem 11 dígitos
    if (cpf.length !== 11) return false;
    
    // Verificar se todos os dígitos são iguais
    if (/^(\d)\1{10}$/.test(cpf)) return false;
    
    // Validar primeiro dígito verificador
    let sum = 0;
    for (let i = 0; i < 9; i++) {
        sum += parseInt(cpf.charAt(i)) * (10 - i);
    }
    let remainder = (sum * 10) % 11;
    if (remainder === 10 || remainder === 11) remainder = 0;
    if (remainder !== parseInt(cpf.charAt(9))) return false;
    
    // Validar segundo dígito verificador
    sum = 0;
    for (let i = 0; i < 10; i++) {
        sum += parseInt(cpf.charAt(i)) * (11 - i);
    }
    remainder = (sum * 10) % 11;
    if (remainder === 10 || remainder === 11) remainder = 0;
    if (remainder !== parseInt(cpf.charAt(10))) return false;
    
    return true;
}

// Função para contar dependentes órfãos (sem titular correspondente)
function countOrphanDependents(operadoraData) {
    if (!operadoraData || operadoraData.length < 2) return 0;
    
    const matriculaIndex = 6; // Índice da coluna Matrícula
    const grauParentescoIndex = 14; // Índice da coluna Grau de Parentesco
    const cpfIndex = 12; // Índice da coluna CPF
    
    let orphanCount = 0;
    const titulares = new Set();
    const dependentes = [];
    
    // Primeiro, identificar todos os titulares
    for (let i = 1; i < operadoraData.length; i++) {
        const row = operadoraData[i];
        const grauParentesco = row[grauParentescoIndex];
        const matricula = row[matriculaIndex];
        const cpf = row[cpfIndex];
        
        // Se é titular, adicionar à lista de titulares
        if (grauParentesco && grauParentesco.toLowerCase().includes('titular')) {
            titulares.add(matricula);
        }
        // Se é dependente, armazenar para verificação posterior
        else if (grauParentesco && !grauParentesco.toLowerCase().includes('titular')) {
            dependentes.push({
                matricula: matricula,
                cpf: cpf,
                grauParentesco: grauParentesco,
                rowIndex: i
            });
        }
    }
    
    // Verificar dependentes sem titular
    dependentes.forEach(dependente => {
        if (!titulares.has(dependente.matricula)) {
            orphanCount++;
        }
    });
    
    return orphanCount;
}

// Função para verificar se existe um titular para uma matrícula específica
function hasTitularInData(data, matricula, mapping) {
    if (!matricula || !data || data.length < 2) return false;
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowMatricula = row[mapping['matricula']] || '';
        const grauParentesco = row[mapping['grau-parentesco']] || '';
        
        // Se encontrou a mesma matrícula e é titular
        if (rowMatricula === matricula && 
            grauParentesco && 
            grauParentesco.toLowerCase().includes('titular')) {
            return true;
        }
    }
    
    return false;
}

function downloadTemplate() {
    let templateData = [];
    let fileName = 'template_excel_converter.xlsx';
    
    if (currentOperationType === 'titular') {
        templateData = [
            ['Operadora', 'Nº da apólice/contrato', 'Nº Sub', 'Inicio de Vigência', 'Início de Vigência Programada', 'Código/Nome do Plano', 'Matrícula', 'Código Dependente', 'Data de admissão', 'Nome Segurado', 'Sexo', 'Data de Nascimento', 'Nº CPF', 'Estado Civil', 'Grau de Parentesco', 'Peso', 'Altura', 'Data da Tutela', 'Data do Documento', 'Data do Casamento', 'Nome da Mãe', 'Nome do Pai', 'PIS', 'RG', 'Órgão Emissor', 'Data Expedição', 'Nº Centro de Custo', 'Nome Centro de Custo', 'Cargo', 'Tipo de Logradouro', 'Nome do Logradouro', 'Número', 'Complemento', 'Bairro', 'Cidade', 'Estado', 'CEP', 'Cartão nacional de saúde', 'Banco', 'Agência', 'Digito da Agência', 'Conta Corrente', 'Dígito conta corrente', 'Número de nascido vivo', 'Telefone', 'E-mail', 'Lote ou Chamado', 'Documento', 'Sequencia', 'Setor', 'Lotação', 'Local', 'Unidade de Atendimento', 'Unidade de Negociação'],
            ['UNIMED', '123456', '001', '01/01/2024', '01/01/2024', 'PLANO BÁSICO', 'MAT001', '', '15/03/2020', 'João Silva', 'M', '15/03/1985', '123.456.789-00', 'Casado', 'Titular', '75', '1.75', '', '15/03/2020', '10/06/2010', 'Maria Silva', 'José Silva', '12345678901', '12.345.678-9', 'SSP', '15/03/2000', '001', 'ADMINISTRATIVO', 'Analista', 'Rua', 'das Flores', '123', 'Apto 45', 'Centro', 'São Paulo', 'SP', '01234-567', '123456789012345', '001', '1234', '1', '12345-6', '7', '123456789012345', '(11) 99999-9999', 'joao@email.com', 'LOTE001', 'DOC001', '001', 'TI', 'SEDE', 'SÃO PAULO', 'UNIDADE SP', 'NEGOCIAÇÃO SP'],
            ['UNIMED', '123456', '001', '01/01/2024', '01/01/2024', 'PLANO BÁSICO', 'MAT001', 'DEP001', '15/03/2020', 'Maria Silva Filha', 'F', '10/05/2010', '123.456.789-01', 'Solteira', 'Filha', '35', '1.40', '', '15/03/2020', '', 'Maria Silva', 'João Silva', '', '12.345.678-0', 'SSP', '10/05/2010', '001', 'ADMINISTRATIVO', 'Estudante', 'Rua', 'das Flores', '123', 'Apto 45', 'Centro', 'São Paulo', 'SP', '01234-567', '123456789012346', '001', '1234', '1', '12345-6', '7', '123456789012346', '(11) 99999-9998', 'maria.filha@email.com', 'LOTE001', 'DOC001', '002', 'TI', 'SEDE', 'SÃO PAULO', 'UNIDADE SP', 'NEGOCIAÇÃO SP'],
            ['SULAMÉRICA', '789012', '002', '01/02/2024', '01/02/2024', 'PLANO COMPLETO', 'MAT002', '', '22/07/2019', 'Maria Santos', 'F', '22/07/1990', '987.654.321-00', 'Solteira', 'Titular', '60', '1.65', '', '22/07/2019', '', 'Ana Santos', 'Pedro Santos', '98765432109', '98.765.432-1', 'SSP', '22/07/2005', '002', 'FINANCEIRO', 'Contador', 'Avenida', 'Paulista', '456', 'Sala 10', 'Bela Vista', 'São Paulo', 'SP', '01310-000', '987654321098765', '341', '5678', '2', '98765-4', '3', '987654321098765', '(11) 88888-8888', 'maria@email.com', 'LOTE002', 'DOC002', '002', 'FINANCEIRO', 'SEDE', 'SÃO PAULO', 'UNIDADE SP', 'NEGOCIAÇÃO SP'],
            ['AMIL', '345678', '003', '01/03/2024', '01/03/2024', 'PLANO PREMIUM', 'MAT003', '', '10/01/2021', 'Carlos Oliveira', 'M', '05/12/1978', '456.789.123-00', 'Casado', 'Titular', '80', '1.80', '', '10/01/2021', '15/08/2005', 'Lucia Oliveira', 'Roberto Oliveira', '34567890123', '34.567.890-1', 'SSP', '05/12/1995', '003', 'VENDAS', 'Vendedor', 'Rua', 'das Palmeiras', '789', 'Casa', 'Jardim', 'Rio de Janeiro', 'RJ', '20000-000', '345678901234567', '237', '9012', '3', '34567-8', '9', '345678901234567', '(21) 77777-7777', 'carlos@email.com', 'LOTE003', 'DOC003', '003', 'VENDAS', 'FILIAL', 'RIO DE JANEIRO', 'UNIDADE RJ', 'NEGOCIAÇÃO RJ'],
            ['UNIMED', '123457', '004', '01/04/2024', '01/04/2024', 'PLANO BÁSICO', 'MAT004', '', '20/05/2022', 'Ana Costa', 'F', '20/05/1992', '111.111.111-11', 'Solteira', 'Titular', '55', '1.60', '', '20/05/2022', '', 'Lucia Costa', 'Paulo Costa', '11111111111', '11.111.111-1', 'SSP', '20/05/2010', '004', 'RH', 'Assistente', 'Rua', 'das Margaridas', '321', 'Apto 12', 'Vila Nova', 'São Paulo', 'SP', '04567-890', '111111111111111', '104', '4321', '4', '11111-1', '1', '111111111111111', '(11) 77777-7777', 'ana@email.com', 'LOTE004', 'DOC004', '004', 'RH', 'SEDE', 'SÃO PAULO', 'UNIDADE SP', 'NEGOCIAÇÃO SP'],
            ['SULAMÉRICA', '789013', '005', '01/05/2024', '01/05/2024', 'PLANO COMPLETO', 'MAT005', '', '15/08/2023', 'Pedro Lima', 'M', '15/08/1988', '000.000.000-00', 'Casado', 'Titular', '85', '1.85', '', '15/08/2023', '12/12/2015', 'Rosa Lima', 'Antonio Lima', '00000000000', '00.000.000-0', 'SSP', '15/08/2008', '005', 'TI', 'Desenvolvedor', 'Avenida', 'Brasil', '654', 'Sala 5', 'Centro', 'São Paulo', 'SP', '01234-567', '000000000000000', '033', '8765', '5', '00000-0', '0', '000000000000000', '(11) 66666-6666', 'pedro@email.com', 'LOTE005', 'DOC005', '005', 'TI', 'SEDE', 'SÃO PAULO', 'UNIDADE SP', 'NEGOCIAÇÃO SP'],
            ['AMIL', '345679', '006', '01/06/2024', '01/06/2024', 'PLANO PREMIUM', 'MAT006', 'DEP002', '10/01/2021', 'Filho Órfão', 'M', '15/03/2015', '555.555.555-55', 'Solteiro', 'Filho', '25', '1.20', '', '10/01/2021', '', 'Mãe Desconhecida', 'Pai Desconhecido', '', '55.555.555-5', 'SSP', '15/03/2015', '006', 'VENDAS', 'Estudante', 'Rua', 'Sem Titular', '999', 'Casa', 'Jardim', 'Rio de Janeiro', 'RJ', '20000-001', '555555555555555', '237', '9999', '9', '55555-5', '5', '555555555555555', '(21) 55555-5555', 'filho.orfao@email.com', 'LOTE006', 'DOC006', '006', 'VENDAS', 'FILIAL', 'RIO DE JANEIRO', 'UNIDADE RJ', 'NEGOCIAÇÃO RJ']
        ];
        fileName = 'template_titular.xlsx';
    } else if (currentOperationType === 'dependente') {
        templateData = [
            ['Operadora', 'Nº da apólice/contrato', 'Nº Sub', 'Matrícula (Titular)', 'Certificado', 'Código Dependente', 'Nome do Titular', 'CPF Titular', 'Nome Dependente', 'CPF do Dependente', 'Data de Nascimento', 'Sexo', 'Estado Civil', 'Grau de Parentesco', 'Peso', 'Altura', 'Data da Tutela', 'Data do Documento', 'Data do Casamento', 'Inicio de Vigência', 'Início de Vigência Programada', 'Nome da Mãe do dependente', 'Nome do Pai', 'Número de nascido vivo', 'Cartão nacional de saúde (CNS)', 'Lote ou Chamado', 'Documento', 'Sequencia', 'Setor', 'Lotação', 'Local'],
            ['UNIMED', '123456', '001', 'MAT001', 'CERT001', 'DEP001', 'João Silva', '123.456.789-00', 'Maria Silva Filha', '123.456.789-01', '10/05/2010', 'F', 'Solteira', 'Filha', '35', '1.40', '', '10/05/2010', '', '01/01/2024', '01/01/2024', 'Maria Silva', 'João Silva', '123456789012345', '123456789012345', 'LOTE001', 'DOC001', '001', 'TI', 'SEDE', 'SÃO PAULO'],
            ['SULAMÉRICA', '789012', '002', 'MAT002', 'CERT002', 'DEP002', 'Maria Santos', '987.654.321-00', 'Pedro Santos Filho', '987.654.321-01', '15/08/2012', 'M', 'Solteiro', 'Filho', '40', '1.50', '', '15/08/2012', '', '01/02/2024', '01/02/2024', 'Ana Santos', 'Pedro Santos', '987654321098765', '987654321098765', 'LOTE002', 'DOC002', '002', 'FINANCEIRO', 'SEDE', 'SÃO PAULO']
        ];
        fileName = 'template_dependente.xlsx';
    } else if (currentOperationType === 'exclusao') {
        templateData = [
            ['Tipo de Exclusão', 'Nº da apólice/contrato', 'Nº Sub', 'Operadora', 'Matrícula', 'Certificado', 'Código Dependente', 'Nome do Titular', 'CPF Titular', 'Data de Demissão', 'Nome Dependente', 'Data de Cancelamento', 'Data de Cancelamento Programada', 'Motivo', 'Lote ou Chamado', 'Documento', 'Sequencia', 'Setor', 'Lotação', 'Local'],
            ['Titular', '123456', '001', 'UNIMED', 'MAT001', 'CERT001', '', 'João Silva', '123.456.789-00', '15/03/2024', '', '15/03/2024', '15/03/2024', 'Demissão', 'LOTE001', 'DOC001', '001', 'TI', 'SEDE', 'SÃO PAULO'],
            ['Dependente', '789012', '002', 'SULAMÉRICA', 'MAT002', 'CERT002', 'DEP001', 'Maria Santos', '987.654.321-00', '', 'Pedro Santos Filho', '20/03/2024', '20/03/2024', 'Mudança de Plano', 'LOTE002', 'DOC002', '002', 'FINANCEIRO', 'SEDE', 'SÃO PAULO']
        ];
        fileName = 'template_exclusao.xlsx';
    } else {
        templateData = [
            ['Nome', 'Email', 'Telefone', 'Endereço', 'Cidade', 'Estado', 'CEP', 'CPF'],
            ['João Silva', 'joao@email.com', '(11) 99999-9999', 'Rua A, 123', 'São Paulo', 'SP', '01234-567', '123.456.789-00'],
            ['Maria Santos', 'maria@email.com', '(21) 88888-8888', 'Av B, 456', 'Rio de Janeiro', 'RJ', '20000-000', '987.654.321-00']
        ];
    }
    
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(templateData);
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
    XLSX.writeFile(workbook, fileName);
    
    showNotification('Template baixado com sucesso!', 'success');
}

// Modal de progresso
function showProgressModal() {
    const modal = document.getElementById('progress-modal');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    
    modal.style.display = 'block';
    
    // Simular progresso
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 30;
        if (progress > 100) progress = 100;
        
        progressFill.style.width = progress + '%';
        
        if (progress < 30) {
            progressText.textContent = 'Lendo arquivo...';
        } else if (progress < 60) {
            progressText.textContent = 'Processando dados...';
        } else if (progress < 90) {
            progressText.textContent = 'Separando por operadora...';
        } else {
            progressText.textContent = 'Finalizando conversão...';
        }
        
        if (progress >= 100) {
            clearInterval(interval);
        }
    }, 200);
}

// Função para detectar operadoras únicas nos dados
function detectOperadoras(data, operadoraColumnIndex) {
    if (!data || data.length < 2 || operadoraColumnIndex === undefined) {
        return [];
    }
    
    const operadoras = new Set();
    
    for (let i = 1; i < data.length; i++) {
        const operadora = data[i][operadoraColumnIndex];
        if (operadora && operadora.trim() !== '') {
            operadoras.add(operadora.trim());
        }
    }
    
    return Array.from(operadoras).sort();
}

function hideProgressModal() {
    document.getElementById('progress-modal').style.display = 'none';
    document.getElementById('progress-fill').style.width = '0%';
}

// Notificações
function showNotification(message, type = 'info') {
    // Criar elemento de notificação
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <i class="fas fa-${getNotificationIcon(type)}"></i>
        <span>${message}</span>
        <button onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    // Adicionar estilos
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${getNotificationColor(type)};
        color: white;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        display: flex;
        align-items: center;
        gap: 10px;
        max-width: 400px;
        animation: slideInRight 0.3s ease;
    `;
    
    // Adicionar ao DOM
    document.body.appendChild(notification);
    
    // Remover automaticamente após 5 segundos
    setTimeout(() => {
        if (notification.parentElement) {
            notification.remove();
        }
    }, 5000);
}

function getNotificationIcon(type) {
    switch (type) {
        case 'success': return 'check-circle';
        case 'error': return 'exclamation-circle';
        case 'warning': return 'exclamation-triangle';
        default: return 'info-circle';
    }
}

function getNotificationColor(type) {
    switch (type) {
        case 'success': return '#38a169';
        case 'error': return '#e53e3e';
        case 'warning': return '#d69e2e';
        default: return '#4299e1';
    }
}

// Adicionar estilos CSS para animação da notificação
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
`;
document.head.appendChild(style); 