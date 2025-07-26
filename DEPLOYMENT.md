# 🚀 Guia para Colocar o Projeto Online

## Opção 1: GitHub Pages (Recomendado - Gratuito)

### Passo 1: Criar conta no GitHub
1. Acesse [github.com](https://github.com)
2. Clique em "Sign up" e crie uma conta gratuita
3. Confirme seu email

### Passo 2: Criar um novo repositório
1. No GitHub, clique no botão "+" no canto superior direito
2. Selecione "New repository"
3. Nome do repositório: `excel-converter-app`
4. Descrição: `Conversor de Excel para Operadoras de Saúde`
5. Deixe público (Public)
6. **NÃO** marque "Add a README file" (já temos um)
7. Clique em "Create repository"

### Passo 3: Fazer upload dos arquivos
1. No repositório criado, clique em "uploading an existing file"
2. Arraste todos os arquivos do projeto:
   - `index.html`
   - `styles.css`
   - `script.js`
   - `README.md`
   - `.gitignore`
3. Clique em "Commit changes"

### Passo 4: Ativar GitHub Pages
1. Vá em "Settings" (aba do repositório)
2. Role para baixo até "Pages"
3. Em "Source", selecione "Deploy from a branch"
4. Em "Branch", selecione "main" e "/(root)"
5. Clique em "Save"
6. Aguarde alguns minutos para o site ficar disponível

### Passo 5: Acessar o site
- Seu site estará disponível em: `https://SEU-USUARIO.github.io/excel-converter-app`

---

## Opção 2: Netlify (Alternativa - Gratuito)

### Passo 1: Criar conta no Netlify
1. Acesse [netlify.com](https://netlify.com)
2. Clique em "Sign up" e crie uma conta
3. Você pode usar sua conta do GitHub

### Passo 2: Fazer deploy
1. No painel do Netlify, clique em "New site from Git"
2. Conecte com GitHub e selecione o repositório
3. Configure:
   - Build command: deixe vazio
   - Publish directory: deixe vazio (ou `/`)
4. Clique em "Deploy site"

### Passo 3: Configurar domínio personalizado (opcional)
1. Em "Site settings" > "Domain management"
2. Clique em "Add custom domain"
3. Configure seu domínio personalizado

---

## Opção 3: Vercel (Alternativa - Gratuito)

### Passo 1: Criar conta no Vercel
1. Acesse [vercel.com](https://vercel.com)
2. Clique em "Sign up" e crie uma conta
3. Conecte com GitHub

### Passo 2: Fazer deploy
1. Clique em "New Project"
2. Importe o repositório do GitHub
3. Configure:
   - Framework Preset: Other
   - Build Command: deixe vazio
   - Output Directory: deixe vazio
4. Clique em "Deploy"

---

## 🔧 Configurações Adicionais

### Personalizar URL
- **GitHub Pages**: `https://SEU-USUARIO.github.io/excel-converter-app`
- **Netlify**: `https://SEU-SITE.netlify.app`
- **Vercel**: `https://SEU-SITE.vercel.app`

### Adicionar domínio personalizado
1. Compre um domínio (ex: GoDaddy, Namecheap)
2. Configure os DNS para apontar para o serviço escolhido
3. Adicione o domínio nas configurações do serviço

### Configurar HTTPS
- Todos os serviços acima já fornecem HTTPS automaticamente

---

## 📱 Testando o Site

Após o deploy, teste:
1. ✅ Upload de arquivos Excel
2. ✅ Seleção de tipo de operação
3. ✅ Mapeamento de colunas
4. ✅ Conversão e download
5. ✅ Responsividade em mobile
6. ✅ Salvar/carregar parâmetros

---

## 🆘 Solução de Problemas

### Site não carrega
- Verifique se todos os arquivos foram enviados
- Aguarde alguns minutos para o deploy
- Verifique se não há erros no console do navegador

### Funcionalidades não funcionam
- Verifique se o JavaScript está habilitado
- Teste em diferentes navegadores
- Verifique se as bibliotecas externas carregam

### Problemas de CSS
- Limpe o cache do navegador
- Verifique se o arquivo CSS foi enviado corretamente

---

## 🎯 Próximos Passos

1. **Compartilhar o link** com sua equipe
2. **Testar em diferentes dispositivos**
3. **Coletar feedback** dos usuários
4. **Fazer melhorias** baseadas no feedback
5. **Considerar domínio personalizado** para uso profissional

---

**🎉 Parabéns! Seu projeto está online e pronto para uso!** 