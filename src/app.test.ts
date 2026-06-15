import { describe, expect, it } from 'vitest';
import request from 'supertest';
import app from './app.js';

describe('AgendaOrg API', () => {
  it('exposes health check', async () => {
    const response = await request(app).get('/api/health');
    expect(response.status).toBe(200);
    expect(response.body.status).toBe('ok');
  });
});
