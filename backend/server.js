const express = require('express');
const cors = require('cors');
const { spawn } = require('child_process');
const path = require('path');
const app = express();
const port = 3000;

app.use(cors());
app.use(express.static(__dirname)); // permite servir arquivos da pasta atual

// Bloqueio de execução simultânea
let processoRodando = false;

app.get('/api/relatorio', (req, res) => {
  const sistema = req.query.sistema;
  const data = req.query.data;

  if (processoRodando) {
    console.warn("[ERRO] Já existe uma requisição em andamento.");
    return res.status(429).json({ erro: "Já existe uma requisição em andamento. Aguarde e tente novamente." });
  }

  console.log(`[LOG] Requisição recebida - sistema: ${sistema}, data: ${data}`);

  const sistemasDisponiveis = ['new', 'c6', 'facta'];

  if (!sistema || !sistemasDisponiveis.includes(sistema)) {
    console.error(`[ERRO] Sistema inválido: ${sistema}`);
    return res.status(400).json({ erro: 'Sistema inválido ou não implementado.' });
  }

  if (!data) {
    console.error('[ERRO] Data não fornecida.');
    return res.status(400).json({ erro: 'Data obrigatória não fornecida.' });
  }

  const scriptPath = `scripts/${sistema}.py`;
  console.log(`[LOG] Executando script: ${scriptPath} com data: ${data}`);

  processoRodando = true;

  const processo = spawn('python', [scriptPath, data], { cwd: __dirname });

  let resultado = '';
  processo.stdout.on('data', chunk => {
    const output = chunk.toString();
    resultado += output;
    console.log(`[PYTHON STDOUT] ${output}`);
  });

  processo.stderr.on('data', chunk => {
    console.error(`[PYTHON STDERR] ${chunk.toString()}`);
  });

  processo.on('close', () => {
    processoRodando = false;
    console.log('[LOG] Script finalizado. Processando resposta...');

    try {
      const resposta = JSON.parse(resultado);

      if (resposta.erro) {
        console.error(`[ERRO DO SCRIPT] ${resposta.erro}`);
        return res.status(500).json({ erro: resposta.erro });
      }

      const caminho = path.join(__dirname, resposta.arquivo);
      console.log(`[LOG] Enviando planilha gerada: ${resposta.arquivo}`);
      res.download(caminho); // força o download

    } catch (err) {
      console.error(`[ERRO AO PARSEAR RETORNO DO SCRIPT] ${err}`);
      console.error(`[CONTEÚDO RECEBIDO]: ${resultado}`);
      res.status(500).json({ erro: 'Erro ao processar retorno do Python.' });
    }
  });
});

app.listen(port, () => {
  console.log(`[LOG] Servidor rodando em http://localhost:${port}`);
});
