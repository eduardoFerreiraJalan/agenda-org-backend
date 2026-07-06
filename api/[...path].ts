import type { Request, Response } from 'express';
import mongoose from 'mongoose';
import app from '../src/app.js';

let connectionPromise: Promise<typeof mongoose> | undefined;

function connectDatabase() {
  if (mongoose.connection.readyState === 1) return Promise.resolve(mongoose);

  const uri = process.env.MONGODB_URI;
  if (!uri) throw new Error('MONGODB_URI nao configurada');

  connectionPromise ??= mongoose.connect(uri);
  return connectionPromise;
}

export default async function handler(req: Request, res: Response) {
  try {
    await connectDatabase();
    return app(req, res);
  } catch (error) {
    console.error('Falha ao conectar ao MongoDB', error);
    return res.status(500).json({ message: 'Falha ao conectar ao banco de dados' });
  }
}
