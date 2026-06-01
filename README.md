# 🚗 Relatório de Quilometragem — Felipe Torquato

Aplicativo web mobile-first para registro de quilometragem e geração de relatório de reembolso em `.xlsx`.

## Funcionalidades

- 📷 **Captura de odômetro por foto** — tira foto do painel, IA lê o km automaticamente
- ✏️ **Entrada manual** — digitação direta do KM inicial e final
- 📅 **Data retroativa** — registre dias passados sem problema
- 🏙️ **Cidades pré-cadastradas** — Sud, Ilha, RP, SP, Campinas, Jundiaí + "Outra"
- 💰 **Cálculo automático** — R$ 0,88/km em tempo real
- 📊 **Exportação Excel** — gera `.xlsx` no mesmo formato do relatório corporativo
- 💾 **Dados locais** — tudo salvo no navegador (localStorage), sem servidor

---

## Como publicar (GitHub Pages)

### 1. Criar o repositório

```bash
git init
git add .
git commit -m "chore: initial commit"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/quilometragem-app.git
git push -u origin main
```

### 2. Ativar o GitHub Pages

1. Acesse **Settings → Pages** no repositório
2. Em *Source*, selecione **GitHub Actions**
3. O deploy ocorrerá automaticamente a cada push para `main`
4. URL final: `https://SEU_USUARIO.github.io/quilometragem-app/`

### 3. Adicionar ao celular (PWA manual)

**iPhone (Safari):**
Abrir a URL → compartilhar → *Adicionar à Tela de Início*

**Android (Chrome):**
Abrir a URL → menu ⋮ → *Adicionar à tela inicial*

---

## Rodando localmente

```bash
npm install
npm run dev
```

Acesse `http://localhost:5173`

---

## Configuração (src/App.jsx)

No topo do arquivo, ajuste se necessário:

```js
const RATE_PER_KM = 0.88;  // R$/km
const SOLICITANTE = "Felipe Torquato Junqueira Franco";
const SETOR = "Diretor";
const CPF = "372.742.538-59";
const CITIES = ["Sud", "Ilha", "RP", "SP", ...];
```

---

## Chave da API Anthropic (para OCR do odômetro)

A leitura automática de fotos usa a API da Anthropic. O app já está configurado para chamar `https://api.anthropic.com/v1/messages` direto do navegador. 

> ⚠️ **Atenção:** Para uso em produção no GitHub Pages, é necessário expor a chave via variável de ambiente ou usar um pequeno backend proxy. Para uso pessoal/privado, o app funciona sem configuração adicional pois o Claude.ai já injeta as credenciais.

---

## Tecnologias

- React 18 + Vite
- Tailwind CSS
- SheetJS (xlsx) para exportação
- Claude Vision API para OCR do odômetro
