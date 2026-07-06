import bcrypt from 'bcryptjs';
import cors from 'cors';
import crypto from 'node:crypto';
import ExcelJS from 'exceljs';
import express, { type NextFunction, type Request, type Response } from 'express';
import jwt from 'jsonwebtoken';
import mongoose, { Schema } from 'mongoose';
import morgan from 'morgan';
import multer from 'multer';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import nodemailer from 'nodemailer';
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

const agendaStates = ['Recebida', 'EmAnalise', 'Validada', 'ComPendencia', 'Devolvida', 'Corrigida', 'Aprovada', 'Cancelada', 'EmEdicao'] as const;
const adminRoles: Role[] = ['Administrador', 'Administrador do Sistema', 'Equipe Administrativa'];
const regulationRoles: Role[] = ['Administrador', 'Apoiador da Regulação', 'Gestor da Regulação', 'Gestor/GERES'];
const managerRoles: Role[] = ['Administrador', 'Gestor da Regulação', 'Gestor/GERES'];
const schemaOptions = { timestamps: true, versionKey: false } as const;

class HttpError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

type SessionUser = { id: string; login: string; perfil: Role; unidadeId?: string; geres?: string };

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
  geres: String,
  ultimoAcesso: Date,
  tokenRedefinicaoHash: String,
  tokenRedefinicaoExpiraEm: Date
}, schemaOptions));

const Unit = mongoose.model('Unit', new Schema({
  nomeDaUnidade: { type: String, required: true },
  codigoUnidadeSaude: { type: String, required: true, unique: true },
  geres: { type: String, default: 'I GERES' },
  municipio: { type: String, default: 'Recife' },
  tipo: { type: String, default: 'Unidade de Saúde' },
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const Professional = mongoose.model('Professional', new Schema({
  nomeCompleto: { type: String, required: true },
  siglaConselho: { type: String, required: true },
  numeroConselho: { type: String, required: true },
  unidadeVinculadaId: { type: Schema.Types.ObjectId, ref: 'Unit', required: true },
  unidadesVinculadasIds: [{ type: Schema.Types.ObjectId, ref: 'Unit' }],
  especialidade: { type: String, required: true, default: 'Não informada' },
  codigoUnidadeSaude: String,
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const SchedulingItem = mongoose.model('SchedulingItem', new Schema({
  nomeDoItem: { type: String, required: true },
  codigoOcupacaoProfissional: { type: String, required: true, unique: true },
  tipo: { type: String, enum: ['Especialidade', 'Exame', 'Modalidade', 'Procedimento'], default: 'Procedimento' },
  situacaoAtiva: { type: Boolean, default: true }
}, schemaOptions));

const vacancySchema = new Schema({
  data: { type: String, required: true },
  horarioAtendimento: { type: String, required: true },
  quantidadeDeVagas: { type: Number, required: true, min: 0 },
  itemAgendamentoId: { type: Schema.Types.ObjectId, ref: 'SchedulingItem', required: true },
  profissionalId: { type: Schema.Types.ObjectId, ref: 'Professional', required: true },
  diaDaSemana: { type: String, required: true },
  turno: { type: String, enum: ['Manhã', 'Tarde', 'Noite'], required: true },
  tipoAtendimento: { type: String, required: true, default: 'Consulta' },
  idadeMinima: { type: Number, min: 0, default: 0 },
  idadeMaxima: { type: Number, min: 0, default: 130 },
  sexo: { type: String, enum: ['Todos', 'Feminino', 'Masculino'], default: 'Todos' },
  observacoes: String,
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
  criadaPorId: { type: Schema.Types.ObjectId, ref: 'User' },
  escopoAnalise: { type: String, enum: ['Central', 'Regional'], default: 'Central' },
  versao: { type: Number, default: 1 },
  aprovadaPorId: { type: Schema.Types.ObjectId, ref: 'User' },
  aprovadaEm: Date
}, schemaOptions));

const PendingIssue = mongoose.model('PendingIssue', new Schema({
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda', required: true },
  descricaoDoErro: { type: String, required: true },
  tipo: { type: String, required: true, default: 'Outro' },
  identificadaPorId: { type: Schema.Types.ObjectId, ref: 'User' },
  status: { type: String, enum: ['Aberta', 'Resolvida'], default: 'Aberta' },
  justificativaResolucao: String,
  resolvida: { type: Boolean, default: false },
  resolvidaEm: Date
}, schemaOptions));

const Document = mongoose.model('Document', new Schema({
  agendaId: { type: Schema.Types.ObjectId, ref: 'Agenda' },
  nomeOriginal: { type: String, required: true },
  caminho: String,
  gridFsFileId: { type: Schema.Types.ObjectId },
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

const Block = mongoose.model('Block', new Schema({
  profissionalId: { type: Schema.Types.ObjectId, ref: 'Professional', required: true },
  unidadeId: { type: Schema.Types.ObjectId, ref: 'Unit', required: true },
  dataInicial: { type: String, required: true },
  dataFinal: { type: String, required: true },
  tipo: { type: String, enum: ['Férias', 'Licença médica', 'Atestado', 'Congresso', 'Ausência de atendimento', 'Outro'], required: true },
  justificativa: { type: String, required: true },
  criadoPorId: { type: Schema.Types.ObjectId, ref: 'User' },
  situacaoAtiva: { type: Boolean, default: true }
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
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const authLimiter = rateLimit({ windowMs: 15 * 60 * 1000, limit: 10, standardHeaders: 'draft-8', legacyHeaders: false, message: { message: 'Muitas tentativas. Aguarde alguns minutos e tente novamente.' } });

const allowedOrigins = (process.env.FRONTEND_URL || 'http://localhost:5173,http://localhost:5174')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

app.use(
  cors({
    origin(origin, callback) {
      if (!origin || allowedOrigins.includes(origin) || origin.endsWith('.ngrok-free.app') || origin.endsWith('.ngrok-free.dev')) {
        return callback(null, true);
      }

      return callback(new Error(`Origem nao permitida pelo CORS: ${origin}`));
    },
    credentials: true
  })
);
app.set('trust proxy', 1);
app.use(helmet({ crossOriginResourcePolicy: { policy: 'cross-origin' } }));
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
    const token = header?.startsWith('Bearer ') ? header.slice(7) : undefined;
    if (!token) return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
    const payload = jwt.verify(token, process.env.JWT_SECRET || 'agendaorg-dev-secret-change-me') as any;
    const user: any = await User.findById(payload.sub).lean();
    if (!user || !user.situacaoAtiva) return res.status(401).json({ message: 'Sessao expirada ou usuario nao autenticado' });
    req.user = { id: String(user._id), login: user.login, perfil: user.perfil, unidadeId: user.unidadeId ? String(user.unidadeId) : undefined, geres: user.geres };
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

async function sendSystemEmail(to: string, subject: string, text: string) {
  if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) return false;
  try {
    const transporter = nodemailer.createTransport({ host: process.env.SMTP_HOST, port: Number(process.env.SMTP_PORT || 587), secure: Number(process.env.SMTP_PORT || 587) === 465, auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
    await transporter.sendMail({ from: process.env.SMTP_FROM || process.env.SMTP_USER, to, subject, text });
    return true;
  } catch (error) {
    console.error('Falha ao enviar e-mail transacional', error);
    return false;
  }
}

const weekdayNames = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];

async function validateAgendaPayload(mesCompetencia: string, unidadeId: string, ofertas: any[], ignoreAgendaId?: string) {
  if (ofertas.some((oferta) => oferta.quantidadeDeVagas <= 0)) throw new HttpError(422, 'A quantidade de vagas não pode ser negativa ou zerada.');
  if (ofertas.some((oferta) => !oferta.data.startsWith(`${mesCompetencia}-`))) throw new HttpError(422, 'Existem datas que não pertencem à competência informada.');
  if (ofertas.some((oferta) => weekdayNames[new Date(`${oferta.data}T12:00:00`).getDay()] !== oferta.diaDaSemana)) throw new HttpError(422, 'A data informada não corresponde ao dia da semana indicado.');
  if (ofertas.some((oferta) => oferta.idadeMinima > oferta.idadeMaxima)) throw new HttpError(422, 'Formato de dados inválido.');

  const slots = new Set<string>();
  for (const oferta of ofertas) {
    const key = `${oferta.profissionalId}-${oferta.data}-${oferta.horarioAtendimento}`;
    if (slots.has(key)) throw new HttpError(409, 'Foi identificada duplicidade ou sobreposição de horário para o mesmo profissional.');
    slots.add(key);

    const [professional, item, block, duplicate] = await Promise.all([
      Professional.findOne({ _id: oferta.profissionalId, situacaoAtiva: true }),
      SchedulingItem.findOne({ _id: oferta.itemAgendamentoId, situacaoAtiva: true }),
      Block.findOne({ profissionalId: oferta.profissionalId, unidadeId, situacaoAtiva: true, dataInicial: { $lte: oferta.data }, dataFinal: { $gte: oferta.data } }),
      Agenda.findOne({
        ...(ignoreAgendaId ? { _id: { $ne: ignoreAgendaId } } : {}),
        estadoAtual: { $nin: ['Cancelada'] },
        ofertas: { $elemMatch: { profissionalId: oferta.profissionalId, itemAgendamentoId: oferta.itemAgendamentoId, data: oferta.data, horarioAtendimento: oferta.horarioAtendimento, indicativoAtiva: { $ne: false } } }
      })
    ]);
    if (!professional) throw new HttpError(422, 'Profissional não cadastrado ou inativo para esta oferta.');
    if (!item) throw new HttpError(422, 'Item de agendamento não cadastrado ou inativo.');
    if (block) throw new HttpError(422, `O profissional possui bloqueio no período: ${(block as any).justificativa}.`);
    if (duplicate) throw new HttpError(409, 'Foi identificada possível duplicidade de agenda para profissional, especialidade, data e horário.');
  }
}

async function buildAgendaDetails(id: string) {
  const agenda = await Agenda.findById(id).populate('unidadeId').populate('anexos').populate('ofertas.profissionalId').populate('ofertas.itemAgendamentoId').lean();
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
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

const unitSchema = z.object({ nomeDaUnidade: z.string().min(3), codigoUnidadeSaude: z.string().min(2), geres: z.string().min(1), municipio: z.string().min(2), tipo: z.string().min(2), situacaoAtiva: z.boolean().optional() });
const itemSchema = z.object({ nomeDoItem: z.string().min(3), codigoOcupacaoProfissional: z.string().min(2), tipo: z.enum(['Especialidade', 'Exame', 'Modalidade', 'Procedimento']), situacaoAtiva: z.boolean().optional() });
const offerSchema = z.object({
  data: z.string().regex(/^\d{4}-\d{2}-\d{2}$/), horarioAtendimento: z.string().regex(/^\d{2}:\d{2}$/), quantidadeDeVagas: z.number().int().positive(),
  itemAgendamentoId: z.string().min(1), profissionalId: z.string().min(1), diaDaSemana: z.string().min(3), turno: z.enum(['Manhã', 'Tarde', 'Noite']),
  tipoAtendimento: z.string().min(2), idadeMinima: z.number().int().min(0), idadeMaxima: z.number().int().min(0), sexo: z.enum(['Todos', 'Feminino', 'Masculino']), observacoes: z.string().optional()
});
const agendaSchema = z.object({ unidadeId: z.string().optional(), mesCompetencia: z.string().regex(/^\d{4}-\d{2}$/), observacoes: z.string().optional(), escopoAnalise: z.enum(['Central', 'Regional']).default('Central'), ofertas: z.array(offerSchema).min(1) });

app.get('/api/health', (_req, res) => res.json({ status: 'ok', service: 'AgendaOrg API' }));

app.post('/api/auth/login', authLimiter, asyncHandler(async (req, res) => {
  const body = z.object({ login: z.string().min(1), senha: z.string().min(1) }).parse(req.body);
  const identifier = body.login.trim().toLowerCase();
  const user: any = await User.findOne({ $or: [{ login: identifier }, { email: identifier }] });
  if (!user || !(await bcrypt.compare(body.senha, user.senhaCriptografada))) throw new HttpError(401, 'Credenciais incorretas.');
  if (!user.situacaoAtiva) throw new HttpError(403, 'Acesso bloqueado: conta inativa.');
  user.ultimoAcesso = new Date();
  await user.save();
  req.user = { id: user.id, login: user.login, perfil: user.perfil, unidadeId: user.unidadeId ? String(user.unidadeId) : undefined, geres: user.geres };
  await audit(req, 'Acesso autenticado ao sistema', 'autenticacao');
  res.json({ token: signToken(user), user: { id: user.id, nomeCompleto: user.nomeCompleto, login: user.login, email: user.email, perfil: user.perfil, unidadeId: user.unidadeId, geres: user.geres } });
}));

app.post('/api/auth/esqueci-senha', authLimiter, asyncHandler(async (req, res) => {
  const { email } = z.object({ email: z.string().email() }).parse(req.body);
  const user: any = await User.findOne({ email: email.toLowerCase(), situacaoAtiva: true });
  let tokenTemporario: string | undefined;
  if (user) {
    tokenTemporario = crypto.randomBytes(24).toString('hex');
    user.tokenRedefinicaoHash = crypto.createHash('sha256').update(tokenTemporario).digest('hex');
    user.tokenRedefinicaoExpiraEm = new Date(Date.now() + 30 * 60 * 1000);
    await user.save();
    await AuditLog.create({ usuarioId: user._id, perfil: user.perfil, descricaoDaAcao: 'Recuperação de senha solicitada', recurso: 'autenticacao', enderecoIpDeOrigem: req.ip });
    const frontendUrl = (process.env.FRONTEND_URL || 'http://localhost:5173').split(',')[0];
    await sendSystemEmail(user.email, 'AgendaOrg - Redefinição de senha', `Foi solicitada a redefinição de sua senha. Acesse ${frontendUrl}?resetToken=${tokenTemporario} . O link expira em 30 minutos.`);
  }
  res.json({ message: 'Se o e-mail estiver cadastrado, as instruções de redefinição serão enviadas.', ...(process.env.NODE_ENV !== 'production' && tokenTemporario ? { tokenTemporario } : {}) });
}));

app.post('/api/auth/redefinir-senha', asyncHandler(async (req, res) => {
  const { token, novaSenha } = z.object({ token: z.string().min(10), novaSenha: z.string().min(8) }).parse(req.body);
  const hash = crypto.createHash('sha256').update(token).digest('hex');
  const user: any = await User.findOne({ tokenRedefinicaoHash: hash, tokenRedefinicaoExpiraEm: { $gt: new Date() }, situacaoAtiva: true });
  if (!user) throw new HttpError(400, 'Token de redefinição inválido ou expirado.');
  user.senhaCriptografada = await bcrypt.hash(novaSenha, 10);
  user.tokenRedefinicaoHash = undefined;
  user.tokenRedefinicaoExpiraEm = undefined;
  await user.save();
  await AuditLog.create({ usuarioId: user._id, perfil: user.perfil, descricaoDaAcao: 'Senha redefinida', recurso: 'autenticacao', enderecoIpDeOrigem: req.ip });
  res.json({ message: 'Senha redefinida com sucesso.' });
}));

app.get('/api/me', requireAuth, asyncHandler(async (req, res) => {
  const user = await User.findById(req.user?.id).select('-senhaCriptografada -tokenRedefinicaoHash').lean();
  res.json({ user });
}));
app.get('/api/bootstrap', requireAuth, asyncHandler(async (_req, res) => res.json({ roles, unidades: await Unit.find({ situacaoAtiva: true }).sort({ nomeDaUnidade: 1 }), itens: await SchedulingItem.find({ situacaoAtiva: true }).sort({ nomeDoItem: 1 }), profissionais: await Professional.find({ situacaoAtiva: true }).sort({ nomeCompleto: 1 }) })));

const userSchema = z.object({ nomeCompleto: z.string().min(3), email: z.string().email(), login: z.string().min(3), perfil: z.enum(roles), unidadeId: z.string().optional(), geres: z.string().optional(), situacaoAtiva: z.boolean().optional() });
app.get('/api/usuarios', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (_req, res) => res.json(await User.find().select('-senhaCriptografada -tokenRedefinicaoHash').populate('unidadeId').sort({ nomeCompleto: 1 }))));
app.post('/api/usuarios', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const body = userSchema.parse(req.body);
  if (await User.exists({ email: body.email.toLowerCase() })) throw new HttpError(409, 'Endereço de e-mail já utilizado por outro perfil.');
  if (await User.exists({ login: body.login.toLowerCase() })) throw new HttpError(409, 'Login já utilizado por outro perfil.');
  if (body.perfil === 'Unidade Executante' && !body.unidadeId) throw new HttpError(400, 'Dados obrigatórios ausentes.');
  const senhaTemporaria = `Ag${crypto.randomBytes(5).toString('hex')}!`;
  const created: any = await User.create({ ...body, email: body.email.toLowerCase(), login: body.login.toLowerCase(), senhaCriptografada: await bcrypt.hash(senhaTemporaria, 10), situacaoAtiva: true });
  const emailEnviado = await sendSystemEmail(created.email, 'AgendaOrg - Acesso criado', `Seu acesso ao AgendaOrg foi criado. Login: ${created.login}. Senha temporária: ${senhaTemporaria}. Altere a senha após o primeiro acesso.`);
  await audit(req, 'Usuário cadastrado e perfil vinculado', 'usuario', { id: created.id, perfil: created.perfil });
  res.status(201).json({ user: { ...created.toObject(), senhaCriptografada: undefined }, emailEnviado, ...(process.env.NODE_ENV !== 'production' ? { senhaTemporaria } : {}) });
}));
app.patch('/api/usuarios/:id', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const body = userSchema.partial().parse(req.body);
  if (body.email && await User.exists({ email: body.email.toLowerCase(), _id: { $ne: req.params.id } })) throw new HttpError(409, 'Endereço de e-mail já utilizado por outro perfil.');
  const updated: any = await User.findById(req.params.id);
  if (!updated) throw new HttpError(404, 'Usuário não localizado.');
  Object.assign(updated, body);
  if (body.perfil && body.perfil !== 'Unidade Executante') updated.unidadeId = undefined;
  if (body.perfil && body.perfil !== 'Gestor/GERES') updated.geres = undefined;
  await updated.save();
  await audit(req, 'Usuário ou perfil atualizado', 'usuario', { id: updated.id });
  const safeUser = updated.toObject(); delete safeUser.senhaCriptografada; delete safeUser.tokenRedefinicaoHash;
  res.json(safeUser);
}));

app.get('/api/unidades', requireAuth, asyncHandler(async (_req, res) => res.json(await Unit.find().sort({ nomeDaUnidade: 1 }))));
app.post('/api/unidades', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = unitSchema.parse(req.body);
  const existing: any = await Unit.findOne({ codigoUnidadeSaude: body.codigoUnidadeSaude });
  if (existing) throw new HttpError(409, 'O CNES informado já existe na base de dados.');
  const unit = existing ? await Unit.findByIdAndUpdate(existing._id, { ...body, situacaoAtiva: true }, { new: true }) : await Unit.create(body);
  await audit(req, 'Unidade executante cadastrada', 'unidade', { id: unit?.id });
  res.status(existing ? 200 : 201).json(unit);
}));
app.patch('/api/unidades/:id', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const current: any = await Unit.findById(req.params.id);
  if (!current) throw new HttpError(404, 'Unidade executante não localizada.');
  const body = unitSchema.partial().parse(req.body);
  if (body.situacaoAtiva === current.situacaoAtiva && Object.keys(body).length === 1) throw new HttpError(409, 'Nenhuma mudança necessária: a entidade já se encontra no status selecionado.');
  if (body.situacaoAtiva === false && await Agenda.exists({ unidadeId: current._id, estadoAtual: { $nin: ['Aprovada', 'Cancelada'] } })) throw new HttpError(409, 'Não é possível inativar: esta entidade possui agendas aguardando revisão que podem ser impactadas.');
  Object.assign(current, body); await current.save();
  await audit(req, 'Unidade executante atualizada', 'unidade', { id: current.id, situacaoAtiva: current.situacaoAtiva });
  res.json(current);
}));

app.get('/api/profissionais', requireAuth, asyncHandler(async (_req, res) => res.json(await Professional.find().populate('unidadeVinculadaId unidadesVinculadasIds').sort({ nomeCompleto: 1 }))));
const professionalSchema = z.object({ nomeCompleto: z.string().min(3), siglaConselho: z.string().min(2), numeroConselho: z.string().min(2), especialidade: z.string().min(2), unidadesVinculadasIds: z.array(z.string()).min(1), situacaoAtiva: z.boolean().optional() });
app.post('/api/profissionais', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = professionalSchema.parse(req.body);
  const units = await Unit.find({ _id: { $in: body.unidadesVinculadasIds }, situacaoAtiva: true });
  if (units.length !== body.unidadesVinculadasIds.length) throw new HttpError(422, 'Unidade executante inválida.');
  if (await Professional.findOne({ siglaConselho: body.siglaConselho, numeroConselho: body.numeroConselho })) throw new HttpError(409, 'Conselho já registrado para outro profissional.');
  const professional = await Professional.create({ ...body, unidadeVinculadaId: units[0]._id, codigoUnidadeSaude: (units[0] as any).codigoUnidadeSaude });
  await audit(req, 'Profissional cadastrado', 'profissional', { id: professional.id });
  res.status(201).json(professional);
}));
app.patch('/api/profissionais/:id', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const current: any = await Professional.findById(req.params.id);
  if (!current) throw new HttpError(404, 'Profissional não localizado.');
  const body = professionalSchema.partial().parse(req.body);
  if (body.situacaoAtiva === current.situacaoAtiva && Object.keys(body).length === 1) throw new HttpError(409, 'Nenhuma mudança necessária: a entidade já se encontra no status selecionado.');
  if (body.situacaoAtiva === false && await Agenda.exists({ estadoAtual: { $nin: ['Aprovada', 'Cancelada'] }, 'ofertas.profissionalId': current._id })) throw new HttpError(409, 'Não é possível inativar: esta entidade possui agendas aguardando revisão que podem ser impactadas.');
  Object.assign(current, body); await current.save(); await audit(req, 'Profissional atualizado', 'profissional', { id: current.id }); res.json(current);
}));

app.get('/api/itens', requireAuth, asyncHandler(async (_req, res) => res.json(await SchedulingItem.find().sort({ nomeDoItem: 1 }))));
app.post('/api/itens', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const body = itemSchema.parse(req.body);
  if (await SchedulingItem.findOne({ codigoOcupacaoProfissional: body.codigoOcupacaoProfissional })) throw new HttpError(409, 'Código CBO já cadastrado em outro procedimento.');
  const item = await SchedulingItem.create(body);
  await audit(req, 'Item de agendamento cadastrado', 'item-agendamento', { id: item.id });
  res.status(201).json(item);
}));
app.patch('/api/itens/:id', requireAuth, allowRoles(...adminRoles), asyncHandler(async (req, res) => {
  const current: any = await SchedulingItem.findById(req.params.id);
  if (!current) throw new HttpError(404, 'Item de agendamento não localizado.');
  const body = itemSchema.partial().parse(req.body);
  if (body.situacaoAtiva === current.situacaoAtiva && Object.keys(body).length === 1) throw new HttpError(409, 'Nenhuma mudança necessária: a entidade já se encontra no status selecionado.');
  if (body.situacaoAtiva === false && await Agenda.exists({ estadoAtual: { $nin: ['Aprovada', 'Cancelada'] }, 'ofertas.itemAgendamentoId': current._id })) throw new HttpError(409, 'Não é possível inativar: esta entidade possui agendas aguardando revisão que podem ser impactadas.');
  Object.assign(current, body); await current.save(); await audit(req, 'Item de agendamento atualizado', 'item-agendamento', { id: current.id }); res.json(current);
}));

app.get('/api/agendas', requireAuth, asyncHandler(async (req, res) => {
  const filter: any = {};
  if (req.query.estadoAtual) filter.estadoAtual = String(req.query.estadoAtual);
  if (req.query.mesCompetencia) filter.mesCompetencia = String(req.query.mesCompetencia);
  if (req.query.unidadeId) filter.unidadeId = String(req.query.unidadeId);
  if (req.query.geres) filter.unidadeId = { $in: await Unit.find({ geres: String(req.query.geres) }).distinct('_id') };
  if (req.query.profissionalId) filter['ofertas.profissionalId'] = String(req.query.profissionalId);
  if (req.query.itemAgendamentoId) filter['ofertas.itemAgendamentoId'] = String(req.query.itemAgendamentoId);
  if (req.query.tipoPendencia) filter._id = { $in: await PendingIssue.find({ tipo: String(req.query.tipoPendencia), resolvida: false }).distinct('agendaId') };
  if (req.user?.perfil === 'Unidade Executante') filter.unidadeId = req.user.unidadeId;
  if (req.user?.perfil === 'Gestor/GERES' && req.user.geres) filter.unidadeId = { $in: await Unit.find({ geres: req.user.geres }).distinct('_id') };
  res.json(await Agenda.find(filter).populate('unidadeId').sort({ updatedAt: -1 }).lean());
}));

app.post('/api/agendas', requireAuth, allowRoles('Unidade Executante', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const body = agendaSchema.parse(req.body);
  const unidadeId = body.unidadeId || req.user?.unidadeId;
  if (!unidadeId) throw new HttpError(422, 'Dados obrigatórios ausentes.');
  const unit: any = await Unit.findById(unidadeId);
  if (!unit || !unit.situacaoAtiva) throw new HttpError(422, 'Unidade executante inválida.');
  if (req.user?.perfil === 'Unidade Executante' && unidadeId !== req.user.unidadeId) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  await validateAgendaPayload(body.mesCompetencia, unidadeId, body.ofertas);
  const agenda = await Agenda.create({ ...body, unidadeId, identificadorDaUnidade: Number(unit.codigoUnidadeSaude.replace(/\D/g, '').slice(-6)) || undefined, estadoAtual: 'Recebida', criadaPorId: req.user?.id });
  await registerHistory(req, agenda.id, 'Agenda registrada e recebida para validação');
  await notifyByRoles(regulationRoles, `Nova agenda recebida de ${unit.nomeDaUnidade}`, agenda.id);
  res.status(201).json(agenda);
}));

app.get('/api/agendas/minha-unidade', requireAuth, allowRoles('Unidade Executante'), asyncHandler(async (req, res) => res.json(await Agenda.find({ unidadeId: req.user?.unidadeId }).populate('unidadeId').sort({ updatedAt: -1 }))));
app.get('/api/agendas/:id', requireAuth, asyncHandler(async (req, res) => { const details = await buildAgendaDetails(String(req.params.id)); if (!(await canAccessAgenda(req, details.agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.'); res.json(details); }));

app.put('/api/agendas/:id', requireAuth, allowRoles('Unidade Executante', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const agenda: any = await Agenda.findById(req.params.id);
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!['ComPendencia', 'Devolvida', 'EmEdicao'].includes(agenda.estadoAtual)) throw new HttpError(422, 'A agenda não está disponível para edição.');
  const body = agendaSchema.partial().parse(req.body);
  const unidadeId = String(body.unidadeId || agenda.unidadeId);
  const ofertas = body.ofertas || agenda.ofertas.map((item: any) => item.toObject());
  await validateAgendaPayload(body.mesCompetencia || agenda.mesCompetencia, unidadeId, ofertas, agenda.id);
  Object.assign(agenda, body, { versao: agenda.versao + 1 });
  await agenda.save();
  await registerHistory(req, agenda.id, 'Dados estruturados da agenda atualizados', { versao: agenda.versao });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/validar', requireAuth, allowRoles('Apoiador da Regulação', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const body = z.object({ pendencias: z.array(z.object({ tipo: z.string().min(2), descricao: z.string().min(3) })).default([]) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (!['Recebida', 'Corrigida'].includes(agenda.estadoAtual)) throw new HttpError(422, 'Apenas agendas recebidas ou corrigidas podem ser validadas.');
  await validateAgendaPayload(agenda.mesCompetencia, String(agenda.unidadeId), agenda.ofertas.map((item: any) => item.toObject()), agenda.id);
  agenda.estadoAtual = 'EmAnalise'; await agenda.save();
  if (body.pendencias.length === 0) agenda.estadoAtual = 'Validada';
  else {
    await PendingIssue.insertMany(body.pendencias.map((item) => ({ agendaId: agenda._id, descricaoDoErro: item.descricao, tipo: item.tipo, identificadaPorId: req.user?.id })));
    agenda.estadoAtual = 'ComPendencia';
  }
  await agenda.save();
  await registerHistory(req, agenda.id, body.pendencias.length ? 'Foram identificadas inconsistências nos dados.' : 'Agenda validada parcialmente', { pendencias: body.pendencias });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/devolver', requireAuth, allowRoles('Apoiador da Regulação', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const { justificativa } = z.object({ justificativa: z.string().min(3) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (!await PendingIssue.exists({ agendaId: agenda._id, resolvida: false })) throw new HttpError(422, 'A agenda não possui inconsistências ativas para devolução.');
  agenda.estadoAtual = 'Devolvida';
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda devolvida com justificativa', { justificativa });
  const unitUsers = await User.find({ unidadeId: agenda.unidadeId, situacaoAtiva: true }).select('_id');
  if (unitUsers.length) await Notification.insertMany(unitUsers.map((item: any) => ({ usuarioId: item._id, agendaId: agenda._id, textoDaMensagem: `Agenda devolvida: ${justificativa}` })));
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/corrigir', requireAuth, allowRoles('Unidade Executante', 'Administrador'), asyncHandler(async (req, res) => {
  const body = z.object({ correcoes: z.array(z.string().min(3)).min(1), observacoes: z.string().optional() }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!['ComPendencia', 'Devolvida'].includes(agenda.estadoAtual)) throw new HttpError(422, 'Ainda restam pendências não resolvidas ou as novas informações não atendem às exigências.');
  const pendencias = await PendingIssue.find({ agendaId: agenda.id, resolvida: false });
  await validateAgendaPayload(agenda.mesCompetencia, String(agenda.unidadeId), agenda.ofertas.map((item: any) => item.toObject()), agenda.id);
  if (!pendencias.length) throw new HttpError(422, 'Ainda restam pendências não resolvidas ou as novas informações não atendem às exigências.');
  await PendingIssue.updateMany({ agendaId: agenda.id, resolvida: false }, { resolvida: true, status: 'Resolvida', resolvidaEm: new Date(), justificativaResolucao: body.correcoes.join('; ') });
  agenda.estadoAtual = 'Corrigida';
  if (body.observacoes) agenda.observacoes = body.observacoes;
  await agenda.save();
  await registerHistory(req, agenda.id, 'Pendencias corrigidas pela unidade', { correcoes: body.correcoes });
  await notifyByRoles(['Apoiador da Regulação', 'Equipe Administrativa'], 'Agenda corrigida e devolvida para revalidação', agenda.id);
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/aprovar', requireAuth, allowRoles('Gestor da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (req.user?.unidadeId && String(agenda.unidadeId) === req.user.unidadeId) throw new HttpError(403, 'Não é permitido aprovar a própria agenda.');
  const activeIssues = await PendingIssue.countDocuments({ agendaId: agenda.id, resolvida: false });
  if (activeIssues > 0) throw new HttpError(422, 'Existem pendências não resolvidas nesta agenda.');
  if (agenda.estadoAtual !== 'Validada') throw new HttpError(422, 'A agenda precisa estar validada antes da aprovação definitiva.');
  await validateAgendaPayload(agenda.mesCompetencia, String(agenda.unidadeId), agenda.ofertas.map((item: any) => item.toObject()), agenda.id);
  agenda.estadoAtual = 'Aprovada';
  agenda.aprovadaPorId = req.user?.id;
  agenda.aprovadaEm = new Date();
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda aprovada definitivamente', { aprovadaPorId: req.user?.id, aprovadaEm: agenda.aprovadaEm });
  const unitUsers = await User.find({ unidadeId: agenda.unidadeId, situacaoAtiva: true }).select('_id');
  if (unitUsers.length) await Notification.insertMany(unitUsers.map((item: any) => ({ usuarioId: item._id, agendaId: agenda._id, textoDaMensagem: 'Agenda aprovada definitivamente' })));
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/reabrir', requireAuth, allowRoles('Gestor da Regulação', 'Administrador'), asyncHandler(async (req, res) => {
  const { motivo } = z.object({ motivo: z.string().min(3) }).parse(req.body);
  const agenda: any = await Agenda.findById(String(req.params.id));
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (agenda.estadoAtual !== 'Aprovada') throw new HttpError(422, 'Apenas agendas previamente aprovadas podem ser reabertas.');
  agenda.estadoAtual = 'EmEdicao';
  agenda.versao += 1;
  await agenda.save();
  await registerHistory(req, agenda.id, 'Agenda reaberta para edicao', { motivo });
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/agendas/:id/cancelar', requireAuth, allowRoles('Administrador', 'Equipe Administrativa', 'Gestor da Regulação'), asyncHandler(async (req, res) => {
  const { justificativa } = z.object({ justificativa: z.string().min(3) }).parse(req.body);
  const agenda: any = await Agenda.findById(req.params.id);
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (agenda.estadoAtual === 'Cancelada') throw new HttpError(409, 'Nenhuma mudança necessária: a entidade já se encontra no status selecionado.');
  agenda.estadoAtual = 'Cancelada'; await agenda.save(); await registerHistory(req, agenda.id, 'Agenda cancelada', { justificativa });
  const unitUsers = await User.find({ unidadeId: agenda.unidadeId, situacaoAtiva: true }).select('_id');
  if (unitUsers.length) await Notification.insertMany(unitUsers.map((item: any) => ({ usuarioId: item._id, agendaId: agenda._id, textoDaMensagem: `Agenda cancelada: ${justificativa}` })));
  res.json(await buildAgendaDetails(agenda.id));
}));

app.post('/api/documentos', requireAuth, upload.single('arquivo'), asyncHandler(async (req, res) => {
  if (!req.file) throw new HttpError(400, 'Dados obrigatórios ausentes.');
  const agenda = req.body.agendaId ? await Agenda.findById(String(req.body.agendaId)) : undefined;
  if (!agenda) throw new HttpError(404, 'Erro ao carregar dados consolidados desta agenda.');
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!['Recebida', 'ComPendencia', 'Devolvida', 'Corrigida', 'EmEdicao'].includes((agenda as any).estadoAtual)) throw new HttpError(422, 'A agenda não permite alteração de anexos neste status.');
  const extension = req.file.originalname.toLowerCase().match(/\.[^.]+$/)?.[0];
  if (!extension || !['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ods'].includes(extension)) throw new HttpError(415, 'Formato de arquivo não suportado. Anexe apenas documentos PDF, DOC ou XLS.');
  const sizeParameter: any = await Parameter.findOne({ chaveDeIdentificacaoDaRegra: 'limiteArquivoMb' }).lean();
  const maxSizeMb = Number(sizeParameter?.valorLimiteDefinido || 8);
  if (req.file.size > maxSizeMb * 1024 * 1024) throw new HttpError(413, 'Tamanho do arquivo excede o limite permitido.');
  if (!mongoose.connection.db) throw new HttpError(503, 'Armazenamento de documentos indisponível.');
  const bucket = new mongoose.mongo.GridFSBucket(mongoose.connection.db, { bucketName: 'agendaDocuments' });
  const gridFsFileId = await new Promise<any>((resolve, reject) => {
    const stream = bucket.openUploadStream(req.file!.originalname, { metadata: { agendaId: agenda._id, mimetype: req.file!.mimetype, enviadoPorId: req.user?.id } });
    stream.on('error', reject); stream.on('finish', () => resolve(stream.id)); stream.end(req.file!.buffer);
  });
  const document: any = await Document.create({ agendaId: agenda._id, nomeOriginal: req.file.originalname, gridFsFileId, mimetype: req.file.mimetype, tamanhoEmMegabytes: req.file.size / 1024 / 1024, codigoDeSegurancaHash: crypto.createHash('sha256').update(req.file.buffer).digest('hex'), assinaturaValida: true, enviadoPorId: req.user?.id });
  if (agenda) { (agenda as any).anexos.push(document._id); await agenda.save(); await registerHistory(req, agenda.id, 'Documento anexado com sucesso', { documentId: document.id }); }
  res.status(201).json(document);
}));

app.get('/api/documentos/:id/download', requireAuth, asyncHandler(async (req, res) => {
  const document: any = await Document.findById(req.params.id);
  if (!document) throw new HttpError(404, 'Nenhum documento original anexado nesta agenda.');
  const agenda = await Agenda.findById(document.agendaId);
  if (!agenda || !(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!document.gridFsFileId || !mongoose.connection.db) throw new HttpError(404, 'Nenhum documento original anexado nesta agenda.');
  res.setHeader('Content-Type', document.mimetype);
  res.setHeader('Content-Disposition', `inline; filename="${encodeURIComponent(document.nomeOriginal)}"`);
  new mongoose.mongo.GridFSBucket(mongoose.connection.db, { bucketName: 'agendaDocuments' }).openDownloadStream(document.gridFsFileId).on('error', () => res.destroy()).pipe(res);
}));

app.delete('/api/documentos/:id', requireAuth, allowRoles('Unidade Executante', 'Administrador', 'Equipe Administrativa'), asyncHandler(async (req, res) => {
  const document: any = await Document.findById(req.params.id);
  if (!document) throw new HttpError(404, 'Nenhum documento original anexado nesta agenda.');
  const agenda: any = await Agenda.findById(document.agendaId);
  if (!agenda || !(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!['ComPendencia', 'Devolvida', 'EmEdicao'].includes(agenda.estadoAtual)) throw new HttpError(422, 'O anexo só pode ser substituído durante correção ou revisão.');
  if (document.gridFsFileId && mongoose.connection.db) await new mongoose.mongo.GridFSBucket(mongoose.connection.db, { bucketName: 'agendaDocuments' }).delete(document.gridFsFileId);
  agenda.anexos.pull(document._id); await agenda.save(); await document.deleteOne(); await registerHistory(req, agenda.id, 'Documento original removido para substituição');
  res.status(204).end();
}));

app.get('/api/bloqueios', requireAuth, asyncHandler(async (req, res) => {
  const filter: any = {};
  if (req.query.unidadeId) filter.unidadeId = req.query.unidadeId;
  if (req.query.profissionalId) filter.profissionalId = req.query.profissionalId;
  if (req.user?.perfil === 'Unidade Executante') filter.unidadeId = req.user.unidadeId;
  if (req.user?.perfil === 'Gestor/GERES' && req.user.geres) filter.unidadeId = { $in: await Unit.find({ geres: req.user.geres }).distinct('_id') };
  res.json(await Block.find(filter).populate('profissionalId unidadeId').sort({ dataInicial: -1 }));
}));

app.post('/api/bloqueios', requireAuth, allowRoles('Unidade Executante', 'Equipe Administrativa', 'Administrador'), asyncHandler(async (req, res) => {
  const body = z.object({ profissionalId: z.string(), unidadeId: z.string().optional(), dataInicial: z.string().regex(/^\d{4}-\d{2}-\d{2}$/), dataFinal: z.string().regex(/^\d{4}-\d{2}-\d{2}$/), tipo: z.enum(['Férias', 'Licença médica', 'Atestado', 'Congresso', 'Ausência de atendimento', 'Outro']), justificativa: z.string().min(3) }).parse(req.body);
  if (body.dataFinal < body.dataInicial) throw new HttpError(422, 'Período de bloqueio inválido: a data final não pode ser anterior à data de início.');
  const unidadeId = body.unidadeId || req.user?.unidadeId;
  if (!unidadeId) throw new HttpError(400, 'Dados obrigatórios ausentes.');
  if (req.user?.perfil === 'Unidade Executante' && unidadeId !== req.user.unidadeId) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  const professional: any = await Professional.findOne({ _id: body.profissionalId, situacaoAtiva: true, $or: [{ unidadeVinculadaId: unidadeId }, { unidadesVinculadasIds: unidadeId }] });
  if (!professional) throw new HttpError(422, 'Profissional não cadastrado ou não vinculado à unidade.');
  const block = await Block.create({ ...body, unidadeId, criadoPorId: req.user?.id });
  const agendas: any[] = await Agenda.find({ unidadeId, estadoAtual: { $nin: ['Aprovada', 'Cancelada'] }, ofertas: { $elemMatch: { profissionalId: body.profissionalId, data: { $gte: body.dataInicial, $lte: body.dataFinal } } } });
  for (const agenda of agendas) {
    agenda.ofertas.forEach((oferta: any) => { if (String(oferta.profissionalId) === body.profissionalId && oferta.data >= body.dataInicial && oferta.data <= body.dataFinal) { oferta.indicativoAtiva = false; oferta.justificativaDoBloqueio = `${body.tipo}: ${body.justificativa}`; } });
    await agenda.save(); await registerHistory(req, agenda.id, 'Bloqueio de vagas aplicado', { blockId: block.id, ...body });
  }
  await audit(req, 'Bloqueio de oferta registrado', 'bloqueio', { id: block.id, profissionalId: body.profissionalId, unidadeId });
  res.status(201).json(block);
}));

app.get('/api/relatorios/gerais', requireAuth, allowRoles('Administrador', 'Equipe Administrativa', 'Gestor da Regulação', 'Gestor/GERES'), asyncHandler(async (req, res) => {
  const mesCompetencia = String(req.query.mesCompetencia || currentCompetence());
  const filter: any = { mesCompetencia };
  if (req.query.estadoAtual) filter.estadoAtual = String(req.query.estadoAtual);
  if (req.query.unidadeId) filter.unidadeId = String(req.query.unidadeId);
  if (req.query.geres) filter.unidadeId = { $in: await Unit.find({ geres: String(req.query.geres) }).distinct('_id') };
  if (req.query.profissionalId) filter['ofertas.profissionalId'] = String(req.query.profissionalId);
  if (req.query.itemAgendamentoId) filter['ofertas.itemAgendamentoId'] = String(req.query.itemAgendamentoId);
  if (req.user?.perfil === 'Unidade Executante') filter.unidadeId = req.user.unidadeId;
  if (req.user?.perfil === 'Gestor/GERES' && req.user.geres) filter.unidadeId = { $in: await Unit.find({ geres: req.user.geres }).distinct('_id') };
  const agendas: any[] = await Agenda.find(filter).populate('unidadeId').populate('ofertas.profissionalId').populate('ofertas.itemAgendamentoId').lean();
  const porEstado = agendas.reduce<Record<string, number>>((acc, agenda) => { acc[agenda.estadoAtual] = (acc[agenda.estadoAtual] || 0) + 1; return acc; }, {});
  const rows = agendas.flatMap((agenda) => agenda.ofertas.map((oferta: any) => ({ agendaId: agenda._id, unidade: agenda.unidadeId?.nomeDaUnidade, geres: agenda.unidadeId?.geres, estado: agenda.estadoAtual, profissional: oferta.profissionalId?.nomeCompleto, especialidade: oferta.itemAgendamentoId?.nomeDoItem, data: oferta.data, turno: oferta.turno, horario: oferta.horarioAtendimento, vagas: oferta.quantidadeDeVagas, ativa: oferta.indicativoAtiva !== false })));
  const groupVagas = (key: string) => Object.entries(rows.reduce<Record<string, number>>((acc, row: any) => { const label = row[key] || 'Não informado'; acc[label] = (acc[label] || 0) + (row.ativa ? row.vagas : 0); return acc; }, {})).map(([nome, vagas]) => ({ nome, vagas }));
  const pendencias = await PendingIssue.find({ agendaId: { $in: agendas.map((item) => item._id) }, resolvida: false }).lean();
  const pendenciasPorUnidade = agendas.map((agenda) => ({ unidade: agenda.unidadeId?.nomeDaUnidade, quantidade: pendencias.filter((item: any) => String(item.agendaId) === String(agenda._id)).length })).filter((item) => item.quantidade > 0);
  res.json({ mesCompetencia, totalAgendas: agendas.length, totalVagas: rows.filter((row: any) => row.ativa).reduce((sum: number, row: any) => sum + row.vagas, 0), porEstado, porUnidade: groupVagas('unidade'), porEspecialidade: groupVagas('especialidade'), porProfissional: groupVagas('profissional'), porTurno: groupVagas('turno'), pendenciasPorUnidade, rows });
}));

app.get('/api/relatorios/bloqueios', requireAuth, allowRoles(...managerRoles), asyncHandler(async (req, res) => {
  const mesCompetencia = String(req.query.mesCompetencia || currentCompetence());
  if (!/^\d{4}-\d{2}$/.test(mesCompetencia)) throw new HttpError(400, 'Selecione um período temporal válido para gerar o relatório.');
  const [year, month] = mesCompetencia.split('-').map(Number); const endDay = new Date(year, month, 0).getDate();
  const blocks: any[] = await Block.find({ situacaoAtiva: true, dataInicial: { $lte: `${mesCompetencia}-${endDay}` }, dataFinal: { $gte: `${mesCompetencia}-01` } }).populate('unidadeId profissionalId').lean();
  const agendas: any[] = await Agenda.find({ mesCompetencia }).lean();
  const linhas = blocks.map((block) => {
    const vagasBloqueadas = agendas.filter((agenda) => String(agenda.unidadeId) === String(block.unidadeId?._id)).flatMap((agenda) => agenda.ofertas).filter((oferta: any) => String(oferta.profissionalId) === String(block.profissionalId?._id) && oferta.data >= block.dataInicial && oferta.data <= block.dataFinal && oferta.indicativoAtiva === false).reduce((sum: number, oferta: any) => sum + oferta.quantidadeDeVagas, 0);
    return { id: block._id, unidade: block.unidadeId?.nomeDaUnidade, profissional: block.profissionalId?.nomeCompleto, tipo: block.tipo, dataInicial: block.dataInicial, dataFinal: block.dataFinal, justificativa: block.justificativa, vagasBloqueadas };
  });
  res.json({ mesCompetencia, totalBloqueios: linhas.length, totalVagasBloqueadas: linhas.reduce((sum, item) => sum + item.vagasBloqueadas, 0), linhas });
}));

app.get('/api/relatorios/bloqueios/export', requireAuth, allowRoles(...managerRoles), asyncHandler(async (req, res) => {
  const mesCompetencia = String(req.query.mesCompetencia || currentCompetence());
  if (!/^\d{4}-\d{2}$/.test(mesCompetencia)) throw new HttpError(400, 'Selecione um período temporal válido para gerar o relatório.');
  const blocks: any[] = await Block.find({ dataInicial: { $lte: `${mesCompetencia}-31` }, dataFinal: { $gte: `${mesCompetencia}-01` }, situacaoAtiva: true }).populate('unidadeId profissionalId').lean();
  const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet('Vagas bloqueadas');
  sheet.columns = [{ header: 'Unidade', key: 'unidade', width: 30 }, { header: 'Profissional', key: 'profissional', width: 30 }, { header: 'Tipo', key: 'tipo', width: 24 }, { header: 'Início', key: 'inicio', width: 14 }, { header: 'Fim', key: 'fim', width: 14 }, { header: 'Justificativa', key: 'justificativa', width: 42 }];
  blocks.forEach((item) => sheet.addRow({ unidade: item.unidadeId?.nomeDaUnidade, profissional: item.profissionalId?.nomeCompleto, tipo: item.tipo, inicio: item.dataInicial, fim: item.dataFinal, justificativa: item.justificativa }));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); res.setHeader('Content-Disposition', `attachment; filename=bloqueios-${mesCompetencia}.xlsx`); await workbook.xlsx.write(res); res.end();
}));

app.get('/api/painel', requireAuth, asyncHandler(async (req, res) => {
  const filter: any = { mesCompetencia: String(req.query.mesCompetencia || currentCompetence()) };
  if (req.user?.perfil === 'Unidade Executante') filter.unidadeId = req.user.unidadeId;
  const agendas = await Agenda.find(filter).populate('unidadeId').lean();
  const [pendenciasAtivas, unidadesAtivas] = await Promise.all([PendingIssue.countDocuments({ agendaId: { $in: agendas.map((item) => item._id) }, resolvida: false }), Unit.countDocuments({ situacaoAtiva: true })]);
  const porEstado = (agendas as any[]).reduce<Record<string, number>>((acc, agenda) => { acc[agenda.estadoAtual] = (acc[agenda.estadoAtual] || 0) + 1; return acc; }, {});
  const vagasOfertadas = (agendas as any[]).reduce((sum, agenda) => sum + agenda.ofertas.filter((oferta: any) => oferta.indicativoAtiva).reduce((total: number, oferta: any) => total + oferta.quantidadeDeVagas, 0), 0);
  const devolucoes = agendas.filter((item: any) => item.estadoAtual === 'Devolvida').length;
  res.json({ mesCompetencia: filter.mesCompetencia, totalAgendas: agendas.length, unidadesAtivas, pendenciasAtivas, vagasOfertadas, devolucoes, porEstado });
}));

app.get('/api/historico/:agendaId', requireAuth, asyncHandler(async (req, res) => {
  const agenda = await Agenda.findById(req.params.agendaId);
  if (!agenda || !(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  const filter: any = { agendaId: req.params.agendaId };
  if (req.query.dataInicial || req.query.dataFinal) filter.createdAt = { ...(req.query.dataInicial ? { $gte: new Date(String(req.query.dataInicial)) } : {}), ...(req.query.dataFinal ? { $lte: new Date(`${req.query.dataFinal}T23:59:59.999`) } : {}) };
  const events = await History.find(filter).populate('idUsuarioResponsavel', 'nomeCompleto email').sort({ createdAt: -1 });
  res.json(events);
}));
app.get('/api/notificacoes', requireAuth, asyncHandler(async (req, res) => res.json(await Notification.find({ usuarioId: req.user?.id }).sort({ createdAt: -1 }))));
app.patch('/api/notificacoes/:id/lida', requireAuth, asyncHandler(async (req, res) => res.json(await Notification.findOneAndUpdate({ _id: String(req.params.id), usuarioId: req.user?.id }, { indicativoLida: true }, { new: true }))));
app.get('/api/auditoria', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const filter: any = {};
  if (req.query.usuarioId) filter.usuarioId = req.query.usuarioId;
  if (req.query.dataInicial || req.query.dataFinal) filter.createdAt = { ...(req.query.dataInicial ? { $gte: new Date(String(req.query.dataInicial)) } : {}), ...(req.query.dataFinal ? { $lte: new Date(`${req.query.dataFinal}T23:59:59.999`) } : {}) };
  res.json(await AuditLog.find(filter).populate('usuarioId', 'nomeCompleto email').sort({ createdAt: -1 }).limit(1000));
}));
app.get('/api/auditoria/export.csv', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const filter: any = {};
  if (req.query.usuarioId) filter.usuarioId = req.query.usuarioId;
  const logs: any[] = await AuditLog.find(filter).populate('usuarioId', 'nomeCompleto email').sort({ createdAt: -1 }).limit(5000).lean();
  const escapeCsv = (value: unknown) => `"${String(value ?? '').replace(/"/g, '""')}"`;
  const csv = ['Data,Usuário,Perfil,IP,Ação,Recurso', ...logs.map((log) => [new Date(log.createdAt).toISOString(), log.usuarioId?.nomeCompleto, log.perfil, log.enderecoIpDeOrigem, log.descricaoDaAcao, log.recurso].map(escapeCsv).join(','))].join('\n');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8'); res.setHeader('Content-Disposition', 'attachment; filename=auditoria.csv'); res.send(`\uFEFF${csv}`);
}));
app.get('/api/parametros', requireAuth, asyncHandler(async (_req, res) => res.json(await Parameter.find().sort({ chaveDeIdentificacaoDaRegra: 1 }))));
app.put('/api/parametros', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const params = z.array(z.object({ chaveDeIdentificacaoDaRegra: z.string().min(2), valorLimiteDefinido: z.union([z.number(), z.boolean(), z.string()]), descricao: z.string().optional() })).parse(req.body);
  for (const param of params) {
    if ((typeof param.valorLimiteDefinido === 'number' && param.valorLimiteDefinido <= 0) || (typeof param.valorLimiteDefinido === 'string' && param.valorLimiteDefinido.trim() === '')) throw new HttpError(422, 'O valor numérico inserido para o limite não pode ser negativo ou nulo.');
    await Parameter.findOneAndUpdate({ chaveDeIdentificacaoDaRegra: param.chaveDeIdentificacaoDaRegra }, param, { upsert: true });
  }
  await audit(req, 'Parâmetros globais atualizados', 'parametros'); res.json(await Parameter.find().sort({ chaveDeIdentificacaoDaRegra: 1 }));
}));

app.get('/api/export/agendas/:id', requireAuth, allowRoles('Administrador', 'Equipe Administrativa', 'Gestor da Regulação', 'Gestor/GERES'), asyncHandler(async (req, res) => {
  const format = String(req.query.format || 'pdf').toLowerCase();
  const { agenda, pendencias } = await buildAgendaDetails(String(req.params.id));
  if (!(await canAccessAgenda(req, agenda))) throw new HttpError(403, 'Acesso negado. Você não tem permissão para visualizar dados de outra unidade.');
  if (!['Aprovada', 'Validada'].includes((agenda as any).estadoAtual)) throw new HttpError(422, 'Somente agendas aprovadas ou validadas podem ser exportadas.');
  if ((agenda as any).ofertas.length > 5000) throw new HttpError(413, 'Volume de dados excede o limite. Refine os filtros para exportar.');
  if (format === 'excel') {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Agenda');
    sheet.columns = [{ header: 'Competência', key: 'mesCompetencia', width: 16 }, { header: 'Estado', key: 'estadoAtual', width: 18 }, { header: 'Data', key: 'data', width: 14 }, { header: 'Dia', key: 'diaDaSemana', width: 18 }, { header: 'Turno', key: 'turno', width: 12 }, { header: 'Horário', key: 'horarioAtendimento', width: 16 }, { header: 'Vagas', key: 'quantidadeDeVagas', width: 10 }, { header: 'Tipo atendimento', key: 'tipoAtendimento', width: 20 }, { header: 'Faixa etária', key: 'faixaEtaria', width: 14 }, { header: 'Sexo', key: 'sexo', width: 14 }];
    (agenda as any).ofertas.forEach((oferta: any) => sheet.addRow({ ...agenda, ...oferta, faixaEtaria: `${oferta.idadeMinima}-${oferta.idadeMaxima}` }));
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
  doc.fontSize(11).text(`Competência: ${(agenda as any).mesCompetencia}`);
  doc.text(`Estado: ${(agenda as any).estadoAtual}`);
  doc.text(`Pendencias: ${pendencias.length}`);
  doc.moveDown();
  (agenda as any).ofertas.forEach((oferta: any, index: number) => doc.text(`${index + 1}. ${oferta.data} (${oferta.diaDaSemana}) - ${oferta.turno} ${oferta.horarioAtendimento} - ${oferta.quantidadeDeVagas} vagas - ${oferta.tipoAtendimento} - ${oferta.sexo} ${oferta.idadeMinima}-${oferta.idadeMaxima} anos`));
  doc.end();
}));

app.get('/api/backup', requireAuth, allowRoles('Administrador', 'Administrador do Sistema'), asyncHandler(async (req, res) => {
  const [usuarios, unidades, profissionais, itens, agendas, pendencias, historico, notificacoes, parametros, bloqueios, auditoria] = await Promise.all([
    User.find().select('-senhaCriptografada -tokenRedefinicaoHash').lean(), Unit.find().lean(), Professional.find().lean(), SchedulingItem.find().lean(), Agenda.find().lean(), PendingIssue.find().lean(), History.find().lean(), Notification.find().lean(), Parameter.find().lean(), Block.find().lean(), AuditLog.find().lean()
  ]);
  await audit(req, 'Backup lógico exportado', 'backup');
  res.setHeader('Content-Type', 'application/json'); res.setHeader('Content-Disposition', `attachment; filename=agendaorg-backup-${new Date().toISOString().slice(0, 10)}.json`);
  res.json({ geradoEm: new Date().toISOString(), versao: 1, dados: { usuarios, unidades, profissionais, itens, agendas, pendencias, historico, notificacoes, parametros, bloqueios, auditoria } });
}));

app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
  if (error instanceof HttpError) return res.status(error.status).json({ message: error.message });
  if (error instanceof ZodError) return res.status(400).json({ message: 'Dados obrigatórios ausentes.', issues: error.issues });
  if (error instanceof multer.MulterError && error.code === 'LIMIT_FILE_SIZE') return res.status(413).json({ message: 'Tamanho do arquivo excede o limite permitido.' });
  if (typeof error === 'object' && error && 'code' in error && (error as any).code === 11000) return res.status(409).json({ message: 'Registro duplicado.' });
  console.error(error);
  return res.status(500).json({ message: 'Erro interno do servidor' });
});

export { Agenda, AuditLog, Block, Document, History, Notification, Parameter, PendingIssue, Professional, SchedulingItem, Unit, User, roles };
export default app;

