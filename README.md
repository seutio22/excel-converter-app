# Conversor de Excel

Uma aplicação web moderna para converter arquivos Excel seguindo parâmetros personalizáveis de mapeamento de colunas.

## 🚀 Funcionalidades

- **Upload de Arquivos Excel**: Suporte para arquivos .xlsx e .xls
- **Drag & Drop**: Interface intuitiva para arrastar e soltar arquivos
- **Prévia dos Dados**: Visualização dos dados antes da conversão
- **Mapeamento de Colunas**: Configuração personalizada de quais colunas representam cada tipo de dado
- **Conversão Inteligente**: Processamento automático seguindo os parâmetros definidos
- **Download de Template**: Template de exemplo para facilitar o uso
- **Interface Responsiva**: Funciona perfeitamente em desktop e mobile

## 📋 Campos Suportados

A aplicação suporta mapeamento para os seguintes campos:

- **Nome**: Nome completo da pessoa
- **Email**: Endereço de email
- **Telefone**: Número de telefone
- **Endereço**: Endereço completo
- **Cidade**: Nome da cidade
- **Estado**: Sigla do estado
- **CEP**: Código postal
- **CPF**: Número do CPF

## 🛠️ Como Usar

### 1. Configuração de Parâmetros

1. Clique na aba "Parâmetros" no topo da página
2. Para cada campo, selecione qual coluna do seu arquivo Excel corresponde
3. Clique em "Salvar Parâmetros" para guardar a configuração

### 2. Upload e Conversão

1. Volte para a aba "Upload"
2. Arraste e solte seu arquivo Excel ou clique para selecionar
3. Visualize a prévia dos dados carregados
4. Clique em "Converter Arquivo" para processar
5. O arquivo convertido será baixado automaticamente

### 3. Template de Exemplo

- Clique em "Baixar Template" para obter um arquivo de exemplo
- Use o template como base para formatar seus dados

## 🎨 Interface

A aplicação possui uma interface moderna e intuitiva com:

- **Design Responsivo**: Adapta-se a diferentes tamanhos de tela
- **Animações Suaves**: Transições elegantes entre páginas
- **Feedback Visual**: Notificações e indicadores de progresso
- **Tema Moderno**: Gradientes e efeitos visuais atrativos

## 🔧 Tecnologias Utilizadas

- **HTML5**: Estrutura semântica
- **CSS3**: Estilos modernos com Flexbox e Grid
- **JavaScript ES6+**: Lógica da aplicação
- **SheetJS (XLSX)**: Manipulação de arquivos Excel
- **Font Awesome**: Ícones
- **LocalStorage**: Persistência de configurações

## 📱 Compatibilidade

- ✅ Chrome (recomendado)
- ✅ Firefox
- ✅ Safari
- ✅ Edge
- ✅ Dispositivos móveis

## 🚀 Como Executar

1. Baixe todos os arquivos para uma pasta
2. Abra o arquivo `index.html` em qualquer navegador moderno
3. A aplicação funcionará completamente offline

## 📝 Estrutura de Arquivos

```
excel-converter-app/
├── index.html          # Página principal
├── styles.css          # Estilos da aplicação
├── script.js           # Lógica JavaScript
└── README.md           # Este arquivo
```

## 🔒 Privacidade

- Todos os processamentos são feitos localmente no navegador
- Nenhum dado é enviado para servidores externos
- As configurações são salvas apenas no seu navegador (localStorage)

## 🐛 Solução de Problemas

### Arquivo não carrega
- Verifique se o arquivo é um Excel válido (.xlsx ou .xls)
- Certifique-se de que o arquivo não está corrompido

### Conversão não funciona
- Configure os parâmetros de mapeamento primeiro
- Verifique se as colunas selecionadas existem no arquivo

### Interface não carrega corretamente
- Use um navegador atualizado
- Verifique se o JavaScript está habilitado

## 📞 Suporte

Para dúvidas ou problemas, verifique:
1. Se está usando um navegador atualizado
2. Se o arquivo Excel está no formato correto
3. Se os parâmetros foram configurados adequadamente

## 🔄 Versão

**Versão atual**: 1.0.0

---

Desenvolvido com ❤️ para facilitar a conversão de arquivos Excel 