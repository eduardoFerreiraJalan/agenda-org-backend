import bcrypt from 'bcryptjs';
import cors from 'cors';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import express, { type NextFunction, type Request, type Response } from 'express';
import jwt from 'jsonwebtoken';
import mongoose, { Schema } from 'mongoose';
import morgan from 'morgan';
import multer from 'multer';
import PDFDocument from 'pdfkit';
import { z, ZodError } from 'zod';

const roles = [
  'Administrador',
  'Administrador do Sistema',
  'Equipe Administrativa',
  'Unidade Executante',
  'Apoiador da Regulação',
  'Gestor da Regulação',
  'Gestor/GERES',
  'Usuário Autenticado'
] as const;

type Role = (typeof roles)[number];

const agendaStates = ['Recebida', 'Validada', 'ComPendencia', 'Devolvida', 'Corrigida', 'Aprovada', 'EmEdicao'] as const;
const adminRoles: Role[] = ['Administrador', 'Administrador do Sistema', 'Equipe Administrativa'];
const regulationRoles: Role[] = ['Administrador', 'Apoiador da Regulação', 'Gestor da Regulação', 'Gestor/GERES'];
const managerRoles: Role[] = ['Administrador', 'Gestor da Regulação', 'Gestor/GERES'];
const schemaOptions = { timestamps: true, versionKey: false } as const;

class HttpError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

type SessionUser = { id: string; login: string; perfil: Role; unidadeId?: string };

declare global {
  namespace Express {
    interface Request {
      user?: SessionUser;
    }
  }
}

const User = mongoose.model('User', new Schema({
  nomeCompleto: { type: String, required: true },
  login: { type: String, required: true, unique: true, trim: true },
  email: { type: String, required: true, unique: true, lowercase: true },
  senhaCriptografada: { type: String, required: true },
  perfil: { type: String, enum: roles, required: true },
  situacaoAtiva: { type: Boolean, default: true },
  unidadeId: { type: Schema.Types.ObjectId, ref: 'Unit' },
  ultimoAcesso: Date
}, schemaOptions));

const Unit = mongoose.model('Unit', new Schema({
  nomeDaUnidade: { type: String, required: true },
  codigoUnidadeSaude: { type: String, required: true, unique: true },
  geres: { type: String, default: 'I GERES' },
  municipio: { type: String, default: 'Recife' },
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const Professional = mongoose.model('Professional', new Schema({
  nomeCompleto: { type: String, required: true },
  siglaConselho: { type: String, required: true },
  numeroConselho: { type: String, required: true },
  unidadeVinculadaId: { type: Schema.Types.ObjectId, ref: 'Unit', required: true },
  codigoUnidadeSaude: String,
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const SchedulingItem = mongoose.model('SchedulingItem', new Schema({
  nomeDoItem: { type: String, required: true },
  codigoOcupacaoProfissional: { type: String, required: true, unique: true },
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const vacancySchema = new Schema({
  data: { type: String, required: true },
  horarioAtendimento: { type: String, required: true },
  quantidadeDeVagas: { type: Number, required: true, min: 0 },
  itemAgendamentoId: { type: Schema.Types.ObjectId, ref: 'SchedulingItem' },
  profissionalId: { type: Schema.Types.ObjectId, ref: 'Professional' },
  indicativoAtiva: { type: Boolean, default: true },
  justificativaDoBloqueio: String
}, { _id: true });

const Agenda = mongoose.model('Agenda', new Schema({
  unidadeId: { type: Schema.Types.ObjectId, ref: 'Unit', required: true },
  identificadorDaUnidade: Number,
  mesCompetencia: { type: String, required: true },
  estadoAtual: { type: String, enum: agendaStates, default: 'Recebida' },
  observacoes: String,
  ofertas: [vacancySchema],
  anexos: [{ type: Schema.Types.ObjectId, ref: 'Document' }],
  criadaPorId: { type: Schema.Types.ObjectId, ref: 'User' }
}, schemaOptions));

const PendingIssue = mongoose.model('PendingIssue', new Schema({
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda', required: true },
  descricaoDoErro: { type: String, required: true },
  resolvida: { type: Boolean, default: false },
  resolvidaEm: Date
}, schemaOptions));

const Document = mongoose.model('Document', new Schema({
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda' },
  nomeOriginal: { type: String, required: true },
  caminho: { type: String, required: true },
  mimetype: { type: String, required: true },
  tamanhoEmMegabytes: { type: Number, required: true },
  codigoDeSegurancaHash: { type: String, required: true },
  assinaturaValida: { type: Boolean, default: true },
  enviadoPorId: { type: Schema.Types.ObjectId, ref: 'User' }
}, schemaOptions));

const History = mongoose.model('History', new Schema({
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda' },
  descricaoDaAcao: { type: String, required: true },
  idUsuarioResponsavel: { type: Schema.Types.ObjectId, ref: 'User' },
  perfil: String,
  enderecoIpDeOrigem: String,
  metadados: Schema.Types.Mixed
}, schemaOptions));

const Notification = mongoose.model('Notification', new Schema({
  usuarioId: { type: Schema.Types.ObjectId, ref: 'User', required: true },
  textoDaMensagem: { type: String, required: true },
  indicativoLida: { type: Boolean, default: false },
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda' }
}, schemaOptions));

const Parameter = mongoose.model('Parameter', new Schema({
  chaveDeIdentificacaoDaRegra: { type: String, required: true, unique: true },
  valorLimiteDefinido: { type: Schema.Types.Mixed, required: true },
  descricao: String
}, schemaOptions));

const AuditLog = mongoose.model('AuditLog', new Schema({
  usuarioId: { type: Schema.Types.ObjectId, ref: 'User' },
  perfil: String,
  descricaoDaAcao: { type: String, required: true },
  enderecoIpDeOrigem: String,
  recurso: String,
  metadados: Schema.Types.Mixed
}, schemaOptions));

const app = express();
const uploadDir = path.resolve(process.cwd(), process.env.UPLOAD_DIR || 'uploads');
fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({ dest: uploadDir, limits: { fileSize: 8 * 1024 * 1024 } });

const allowedOrigins = (process.env.FRONTEND_URL || 'http://localhost:5173,http://localhost:5174')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

app.use(
  cors({
    origin(origin, callback) {
      if (!origin || allowedOrigins.includes(origin)) {
        return callback(null, true);
      }

      return callback(new Error(`Origem nao permitida pelo CORS: ${origin}`));
    },
    credentials: true
  })
);
app.use(express.json({ limit: '2mb' }));
app.use(morgan('dev'));

function asyncHandler(fn: (req: Request, res: Response, next: NextFunction) => Promise<unknown>) {
  return (req: Request, res: Response, next: NextFunction) => Promise.resolve(fn(req, res, next)).catch(next);
}

function signToken(user: any) {
  return jwt.sign({ sub: String(user._id), login: user.login, role: user.perfil }, process.env.JWT_SECRET || 'agendaorg-dev-secret-change-me', { expiresIn: '8h' });
}

async function requireAuth(req: Request, res: Response, next: NextFunction) {
  try {
    const header = req.headers.authorization;
    const queryToken = typeof req.query.token === 'string' ? req.query.token : undefined;
    const token = header?.startsWith('Bearer ') ? header.slice(7) : queryToken;
    if (!token) return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
    const payload = jwt.verify(token, process.env.JWT_SECRET || 'agendaorg-dev-secret-change-me') as any;
    const user: any = await User.findById(payload.sub).lean();
    if (!user || !user.situacaoAtiva) return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
    req.user = { id: String(user._id), login: user.login, perfil: user.perfil, unidadeId: user.unidadeId ? String(user.unidadeId) : undefined };
    return next();
  } catch {
    return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
  }
}

function allowRoles(...allowed: Role[]) {
  return (req: Request, res: Response, next: NextFunction) => {
    if (!req.user) return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
    if (!allowed.includes(req.user.perfil)) return res.status(403).json({ message: 'Operacao nao autorizada' });
    return next();
  };
}

async function audit(req: Request, descricaoDaAcao: string, recurso?: string, metadados?: unknown) {
  await AuditLog.create({ usuarioId: req.user?.id, perfil: req.user?.perfil, descricaoDaAcao, recurso, enderecoIpDeOrigem: req.ip, metadados });
}

async function registerHistory(req: Request, agendaId: string, descricaoDaAcao: string, metadados?: unknown) {
  await History.create({ agendaId, descricaoDaAcao, idUsuarioResponsavel: req.user?.id, perfil: req.user?.perfil, enderecoIpDeOrigem: req.ip, metadados });
  await audit(req, descricaoDaAcao, 'agenda', { agendaId, metadados });
}

async function notifyByRoles(perfis: string[], textoDaMensagem: string, agendaId?: string) {
  const users = await User.find({ perfil: { $in: perfis }, situacaoAtiva: true }).select('_id').lean();
  if (users.length) await Notification.insertMany(users.map((user: any) => ({ usuarioId: user._id, textoDaMensagem, agendaId })));
}

function currentCompetence() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function validateAgendaPayload(mesCompetencia: string, ofertas: any[]) {
  if (ofertas.some((oferta) => oferta.quantidadeDeVagas <= 0)) throw new HttpError(422, 'Existem datas sem vagas ofertadas');
  if (ofertas.some((oferta) => !oferta.data.startsWith(`${mesCompetencia}-`))) throw new HttpError(422, 'Datas nao pertencem a competencia');
  const slots = new Set<string>();
  for (const oferta of ofertas) {
    const key = `${oferta.data}-${oferta.horarioAtendimento}`;
    if (slots.has(key)) throw new HttpError(422, 'Horario duplicado na agenda');
    slots.add(key);
  }
}

async function buildAgendaDetails(id: string) {
  const agenda = await Agenda.findById(id).populate('unidadeId').populate('anexos').lean();
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  const [pendencias, historico] = await Promise.all([
    PendingIssue.find({ agendaId: id }).sort({ createdAt: -1 }).lean(),
    History.find({ agendaId: id }).sort({ createdAt: -1 }).lean()
  ]);
  return { agenda, pendencias, historico };
}

async function canAccessAgenda(req: Request, agenda: any) {
  if (req.user?.perfil !== 'Unidade Executante') return true;
  return String(agenda.unidadeId?._id || agenda.unidadeId) === req.user.unidadeId;
}

const unitSchema = z.object({ nomeDaUnidade: z.string().min(3), codigoUnidadeSaude: z.string().min(2), geres: z.string().optional(), municipio: z.string().optional(), situacaoAtiva: z.boolean().optional() });
const itemSchema = z.object({ nomeDoItem: z.string().min(3), codigoOcupacaoProfissional: z.string().min(2), situacaoAtiva: z.boolean().optional() });
const agendaSchema = z.object({ unidadeId: z.string().optional(), mesCompetencia: z.string().regex(/^\d{4}-\d{2}$/), observacoes: z.string().optional(), ofertas: z.array(z.object({ data: z.string().regex(/^\d{4}-\d{2}-\d{2}$/), horarioAtendimento: z.string().min(4), quantidadeDeVagas: z.number().int().min(0), itemAgendamentoId: z.string().optional(), profissionalId: z.string().optional() })).min(1) });

app.get('/api/health', (_req, res) => res.json({ status: 'ok', service: 'AgendaOrg API' }));

app.post('/api/auth/login', asyncHandler(async (req, res) => {
  const body = z.object({ login: z.string(), senha: z.string() }).parse(req.body);
  const user: any = await User.findOne({ login: body.login });
  if (!user) throw new HttpError(401, 'Usuario nao encontrado');
  if (!user.situacaoAtiva) throw new HttpError(403, 'Conta inativa');
  if (!(await bcrypt.compare(body.senha, user.senhaCriptografada))) throw new HttpError(401, 'Credenciais invalidas');
  user.ultimoAcesso = new Date();
  await user.save();
  res.json({ token: signToken(user), user: { id: user.id, nomeCompleto: user.nomeCompleto, login: user.login, email: user.email, perfil: user.perfil, unidadeId: user.unidadeId } });
}));

app.get('/api/me', requireAuth, (req, res) => res.json({ user: req.user }));
app.get('/api/bootstrap', requireAuth, asyncHandler(async (_req, res) => res.json({ roles, unidades: await Unit.find().sort({ nomeDaUnidade: 1 }), itens: await SchedulingItem.find().sort({ nomeDoItem: 1 }), profissionais: await Professional.find().sort({ nomeCompleto: 1 }) })));

app.get('/api/unidades', requireAuth, asyncHandler(async (_req, res) => res.json(await Unit.find().sort({ nomeDaUnidade: 1 }))));
app.post('/api/unidades', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = unitSchema.parse(req.body);
  const existing: any = await Unit.findOne({ codigoUnidadeSaude: body.codigoUnidadeSaude });
  if (existing && existing.situacaoAtiva) throw new HttpError(409, 'Unidade ja cadastrada');
  const unit = existing ? await Unit.findByIdAndUpdate(existing._id, { ...body, situacaoAtiva: true }, { new: true }) : await Unit.create(body);
  await audit(req, 'Unidade executante cadastrada', 'unidade', { id: unit?.id });
  res.status(existing ? 200 : 201).json(unit);
}));
app.patch('/api/unidades/:id', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => res.json(await Unit.findByIdAndUpdate(String(req.params.id), unitSchema.partial().parse(req.body), { new: true }))));

app.get('/api/profissionais', requireAuth, asyncHandler(async (_req, res) => res.json(await Professional.find().populate('unidadeVinculadaId').sort({ nomeCompleto: 1 }))));
app.post('/api/profissionais', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = z.object({ nomeCompleto: z.string().min(3), siglaConselho: z.string().min(2), numeroConselho: z.string().min(2), codigoUnidadeSaude: z.string().min(2) }).parse(req.body);
  const unit = await Unit.findOne({ codigoUnidadeSaude: body.codigoUnidadeSaude, situacaoAtiva: true });
  if (!unit) throw new HttpError(422, 'Unidade executante invalida');
  if (await Professional.findOne({ siglaConselho: body.siglaConselho, numeroConselho: body.numeroConselho })) throw new HttpError(409, 'Profissional ja cadastrado');
  const professional = await Professional.create({ ...body, unidadeVinculadaId: unit._id });
  await audit(req, 'Profissional cadastrado', 'profissional', { id: professional.id });
  res.status(201).json(professional);
}));

app.get('/api/itens', requireAuth, asyncHandler(async (_req, res) => res.json(await SchedulingItem.find().sort({ nomeDoItem: 1 }))));
app.post('/api/itens', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = itemSchema.parse(req.body);
  if (await SchedulingItem.findOne({ codigoOcupacaoProfissional: body.codigoOcupacaoProfissional })) throw new HttpError(409, 'Codigo ocupacional ja cadastrado');
  const item = await SchedulingItem.create(body);
  await audit(req, 'Item de agendamento cadastrado', 'item-agendamento', { id: item.id });
  res.status(201).json(item);
}));

app.get('/api/agendas', requireAuth, asyncHandler(async (req, res) => {
  const filter: any = {};
  if (req.query.estadoAtual) filter.estadoAtual = String(req.query.estadoAtual);
  if (req.query.mesCompetencia) filter.mesCompetencia = String(req.query.mesCompetencia);
  if (req.query.unidadeId) filter.unidadeId = String(req.query.unidadeId);
  if (req.user?.perfil === 'Unidade Executante') filter.unidadeId = req.user.unidadeId;
  res.json(await Agenda.find(filter).populate('unidadeId').sort({ updatedAt: -1 }).lean());
}));

app.post('/api/agendas', requireAuth, allowRoles('Unidade Executante', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const body = agendaSchema.parse(req.body);
  const unidadeId = body.unidadeId || req.user?.unidadeId;
  if (!unidadeId) throw new HttpError(422, 'Unidade executante invalida');
  const unit: any = await Unit.findById(unidadeId);
  if (!unit || !unit.situacaoAtiva) throw new HttpError(422, 'Unidade executante invalida');
  validateAgendaPayload(body.mesCompetencia, body.ofertas);
  const agenda = await Agenda.create({ ...body, unidadeId, identificadorDaUnidade: Number(unit.codigoUnidadeSaude.replace(/\D/g, '').slice(-6)) || undefined, estadoAtual: 'Recebida', criadaPorId: req.user?.id });
  await registerHistory(req, agenda.id, 'Agenda registrada e recebida para validação');
  await notifyByRoles(regulationRoles, `Nova agenda recebida de ${unit.nomeDaUnidade}`, agenda.id);
  res.status(201).json(agenda);
}));

app.get('/api/agendas/minha-unidade', requireAuth, allowRoles('Unidade Executante'), asyncHandler(async (req, res) => res.json(await Agenda.find({ unidadeId: req.user?.unidadeId }).populate('unidadeId').sort({ updatedAt: -1 }))));
app.get('/api/agendas/:id', requireAuth, asyncHandler(async (req, res) => { const details = await buildAgendaDetails(String(req.params.id)); if (!(await canAccessAgenda(req, details.agenda))) throw new HttpError(403, 'Voce nao tem permissao para acessar esta agenda'); res.json(details); }));

app.post('/api/agendas/:id/validar', requireAuth, allowRoles('Apoiador da Regulação', 'Administrador', 'Gestor da Regulação'), asyncHandler(async (req, res) => {
  const body = z.object({ erros: z.array(z.string().min(3)).default([]) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  if (body.erros.length === 0) agenda.estadoAtual = 'Validada';
  else {
    await PendingIssue.insertMany(body.erros.map((descricaoDoErro) => ({ agendaId: agenda._id, descricaoDoErro })));
    agenda.estadoAtual = 'ComPendencia';
  }
  await agenda.save();
  await registerHistory(req, agenda.id, body.erros.length ? 'Agenda devolvida com pendencias' : 'Agenda Validada', { erros: body.erros });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/devolver', requireAuth, allowRoles('Apoiador da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const { justificativa } = z.object({ justificativa: z.string().min(3) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  agenda.estadoAtual = 'Devolvida';
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda devolvida com justificativa', { justificativa });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/corrigir', requireAuth, allowRoles('Unidade Executante', 'Administrador'), asyncHandler(async (req, res) => {
  const body = z.object({ correcoes: z.array(z.string()).default([]), observacoes: z.string().optional() }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Voce nao tem permissao para acessar esta agenda');
  const pendencias = await PendingIssue.find({ agendaId: agenda.id, resolvida: false });
  if (pendencias.length > 0 && body.correcoes.length === 0) throw new HttpError(422, 'Novas informacoes nao atendem as exigencias');
  await PendingIssue.updateMany({ agendaId: agenda.id, resolvida: false }, { resolvida: true, resolvidaEm: new Date() });
  agenda.estadoAtual = 'Corrigida';
  if (body.observacoes) agenda.observacoes = body.observacoes;
  await agenda.save();
  await registerHistory(req, agenda.id, 'Pendencias corrigidas pela unidade', { correcoes: body.correcoes });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/aprovar', requireAuth, allowRoles('Gestor da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  const activeIssues = await PendingIssue.countDocuments({ agendaId: agenda.id, resolvida: false });
  if (activeIssues > 0 || !['Validada', 'Corrigida'].includes(agenda.estadoAtual)) throw new HttpError(422, 'Agenda impossibilitada de aprovacao');
  agenda.estadoAtual = 'Aprovada';
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda aprovada');
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/reabrir', requireAuth, allowRoles('Gestor da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const { motivo } = z.object({ motivo: z.string().min(3) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  if (agenda.estadoAtual !== 'Aprovada') throw new HttpError(422, 'Apenas agendas aprovadas podem ser reabertas');
  agenda.estadoAtual = 'EmEdicao';
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda reaberta para edicao', { motivo });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/documentos', requireAuth, upload.single('arquivo'), asyncHandler(async (req, res) => {
  if (!req.file) throw new HttpError(400, 'Arquivo obrigatorio');
  const agenda = req.body.agendaId ? await Agenda.findById(String(req.body.agendaId)) : undefined;
  if (agenda && !(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Voce nao tem permissao para acessar esta agenda');
  const fileBuffer = fs.readFileSync(req.file.path);
  const document: any = await Document.create({ agendaId: agenda?._id, nomeOriginal: req.file.originalname, caminho: req.file.path, mimetype: req.file.mimetype, tamanhoEmMegabytes: req.file.size / 1024 / 1024, codigoDeSegurancaHash: crypto.createHash('sha256').update(fileBuffer).digest('hex'), assinaturaValida: true, enviadoPorId: req.user?.id });
  if (agenda) { (agenda as any).anexos.push(document._id); await agenda.save(); await registerHistory(req, agenda.id, 'Documento anexado com sucesso', { documentId: document.id }); }
  res.status(201).json(document);
}));

app.post('/api/bloqueios', requireAuth, allowRoles('Gestor da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const body = z.object({ agendaId: z.string(), dataInicial: z.string(), dataFinal: z.string(), justificativa: z.string().min(3) }).parse(req.body);
  if (body.dataFinal < body.dataInicial) throw new HttpError(422, 'Periodo de bloqueio invalido');
  const agenda: any = await Agenda.findById(body.agendaId);
  if (!agenda) throw new HttpError(404, 'Agenda nao localizada ou excluida');
  agenda.ofertas.forEach((oferta: any) => { if (oferta.data >= body.dataInicial && oferta.data <= body.dataFinal) { oferta.indicativoAtiva = false; oferta.justificativaDoBloqueio = body.justificativa; } });
  await agenda.save();
  await registerHistory(req, agenda.id, 'Bloqueio de vagas aplicado em lote', body);
  res.json(agenda);
}));

app.get('/api/relatorios/gerais', requireAuth, asyncHandler(async (req, res) => {
  const mesCompetencia = String(req.query.mesCompetencia || currentCompetence());
  const agendas: any[] = await Agenda.find({ mesCompetencia }).populate('unidadeId').lean();
  const porEstado = agendas.reduce<Record<string, number>>((acc, agenda) => { acc[agenda.estadoAtual] = (acc[agenda.estadoAtual] || 0) + 1; return acc; }, {});
  res.json({ mesCompetencia, total: agendas.length, porEstado, agendas });
}));

app.get('/api/relatorios/bloqueios', requireAuth, allowRoles(...managerRoles), asyncHandler(async (req, res) => {
  const mesCompetencia = String(req.query.mesCompetencia || currentCompetence());
  const agendas: any[] = await Agenda.find({ mesCompetencia }).populate('unidadeId').lean();
  const vagas = agendas.flatMap((agenda) => agenda.ofertas.filter((oferta: any) => !oferta.indicativoAtiva).map((oferta: any) => ({ agendaId: agenda._id, unidade: agenda.unidadeId, data: oferta.data, horarioAtendimento: oferta.horarioAtendimento, justificativaDoBloqueio: oferta.justificativaDoBloqueio })));
  res.json({ mesCompetencia, total: vagas.length, vagas });
}));

app.get('/api/painel', requireAuth, asyncHandler(async (_req, res) => {
  const [agendas, pendenciasAtivas, unidadesAtivas] = await Promise.all([Agenda.find().populate('unidadeId').lean(), PendingIssue.countDocuments({ resolvida: false }), Unit.countDocuments({ situacaoAtiva: true })]);
  const porEstado = (agendas as any[]).reduce<Record<string, number>>((acc, agenda) => { acc[agenda.estadoAtual] = (acc[agenda.estadoAtual] || 0) + 1; return acc; }, {});
  const vagasOfertadas = (agendas as any[]).reduce((sum, agenda) => sum + agenda.ofertas.filter((oferta: any) => oferta.indicativoAtiva).reduce((total: number, oferta: any) => total + oferta.quantidadeDeVagas, 0), 0);
  res.json({ totalAgendas: agendas.length, unidadesAtivas, pendenciasAtivas, vagasOfertadas, porEstado });
}));

app.get('/api/historico/:agendaId', requireAuth, asyncHandler(async (req, res) => res.json(await History.find({ agendaId: String(req.params.agendaId) }).sort({ createdAt: -1 }))));
app.get('/api/notificacoes', requireAuth, asyncHandler(async (req, res) => res.json(await Notification.find({ usuarioId: req.user?.id }).sort({ createdAt: -1 }))));
app.patch('/api/notificacoes/:id/lida', requireAuth, asyncHandler(async (req, res) => res.json(await Notification.findOneAndUpdate({ _id: String(req.params.id), usuarioId: req.user?.id }, { indicativoLida: true }, { new: true }))));
app.get('/api/auditoria', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (_req, res) => res.json(await AuditLog.find().sort({ createdAt: -1 }).limit(200))));
app.get('/api/parametros', requireAuth, asyncHandler(async (_req, res) => res.json(await Parameter.find().sort({ chaveDeIdentificacaoDaRegra: 1 }))));
app.put('/api/parametros', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => { for (const param of req.body) await Parameter.findOneAndUpdate({ chaveDeIdentificacaoDaRegra: param.chaveDeIdentificacaoDaRegra }, param, { upsert: true }); await audit(req, 'Parâmetros globais atualizados', 'parametros'); res.json(await Parameter.find().sort({ chaveDeIdentificacaoDaRegra: 1 })); }));

app.get('/api/export/agendas/:id', requireAuth, asyncHandler(async (req, res) => {
  const format = String(req.query.format || 'pdf').toLowerCase();
  const { agenda, pendencias } = await buildAgendaDetails(String(req.params.id));
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Voce nao tem permissao para acessar esta agenda');
  if (format === 'excel') {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Agenda');
    sheet.columns = [{ header: 'Competencia', key: 'mesCompetencia', width: 16 }, { header: 'Estado', key: 'estadoAtual', width: 18 }, { header: 'Data', key: 'data', width: 14 }, { header: 'Horario', key: 'horarioAtendimento', width: 16 }, { header: 'Vagas', key: 'quantidadeDeVagas', width: 10 }];
    (agenda as any).ofertas.forEach((oferta: any) => sheet.addRow({ ...agenda, ...oferta }));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=agenda-${(agenda as any)._id}.xlsx`);
    await workbook.xlsx.write(res);
    return res.end();
  }
  const doc = new PDFDocument({ margin: 48 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename=agenda-${(agenda as any)._id}.pdf`);
  doc.pipe(res);
  doc.fontSize(18).text('AgendaOrg - Agenda Consolidada');
  doc.moveDown();
  doc.fontSize(11).text(`Competencia: ${(agenda as any).mesCompetencia}`);
  doc.text(`Estado: ${(agenda as any).estadoAtual}`);
  doc.text(`Pendencias: ${pendencias.length}`);
  doc.moveDown();
  (agenda as any).ofertas.forEach((oferta: any, index: number) => doc.text(`${index + 1}. ${oferta.data} - ${oferta.horarioAtendimento} - ${oferta.quantidadeDeVagas} vagas`));
  doc.end();
}));

app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
  if (error instanceof HttpError) return res.status(error.status).json({ message: error.message });
  if (error instanceof ZodError) return res.status(400).json({ message: 'Dados invalidos', issues: error.issues });
  if (typeof error === 'object' && error && 'code' in error && (error as any).code === 11000) return res.status(409).json({ message: 'Registro duplicado' });
  console.error(error);
  return res.status(500).json({ message: 'Erro interno do servidor' });
});

export { Agenda, AuditLog, Document, History, Notification, Parameter, PendingIssue, Professional, SchedulingItem, Unit, User, roles };
export default app;

