import 'dotenv/config';
import mongoose from 'mongoose';
import app from './app.js';

const port = Number(process.env.PORT || 3333);
const uri = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/agendaorg';

mongoose.connect(uri).then(() => {
  const server = app.listen(port, () => console.log(`AgendaOrg API rodando em http://localhost:${port}`));

  server.on('error', (error: NodeJS.ErrnoException) => {
    if (error.code === 'EADDRINUSE') {
      console.error(`A porta ${port} ja esta em uso. Encerre o outro backend ou defina PORT em .env.`);
      process.exit(1);
    }

    console.error('Falha ao iniciar a API', error);
    process.exit(1);
  });
}).catch((error) => {
  console.error('Falha ao conectar no MongoDB', error);
  process.exit(1);
});
