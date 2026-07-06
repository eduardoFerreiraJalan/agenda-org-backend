import { describe, expect, it } from 'vitest';
import request from 'supertest';
import app from './app.js';

describe('AgendaOrg API', () => {
  it('exposes health check', async () => {
    const response = await request(app).get('/api/health');
    expect(response.status).toBe(200);
    expect(response.body.status).toBe('ok');
  });

  it('protects agenda data without authentication', async () => {
    const response = await request(app).get('/api/agendas');
    expect(response.status).toBe(401);
  });

  it('protects global audit without authentication', async () => {
    const response = await request(app).get('/api/auditoria');
    expect(response.status).toBe(401);
  });

  it('allows configured frontend preflight', async () => {
    const response = await request(app).options('/api/auth/login').set('Origin', 'http://localhost:5173').set('Access-Control-Request-Method', 'POST');
    expect(response.status).toBe(204);
    expect(response.headers['access-control-allow-origin']).toBe('http://localhost:5173');
  });

  it('validates password recovery email', async () => {
    const response = await request(app).post('/api/auth/esqueci-senha').send({ email: 'invalido' });
    expect(response.status).toBe(400);
    expect(response.body.message).toBe('Dados obrigatórios ausentes.');
  });
});
