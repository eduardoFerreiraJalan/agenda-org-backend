import 'dotenv/config';
import bcrypt from 'bcryptjs';
import mongoose from 'mongoose';
import { Agenda, AuditLog, Document, History, Notification, Parameter, PendingIssue, Professional, SchedulingItem, Unit, User } from './app.js';

async function main() {
  await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/agendaorg');
  await Promise.all([Agenda.deleteMany({}), AuditLog.deleteMany({}), Document.deleteMany({}), History.deleteMany({}), Notification.deleteMany({}), Parameter.deleteMany({}), PendingIssue.deleteMany({}), Professional.deleteMany({}), SchedulingItem.deleteMany({}), Unit.deleteMany({}), User.deleteMany({})]);

  const [upaeRecife, hospitalCaruaru, upaePetrolina] = await Unit.insertMany([
    { nomeDaUnidade: 'UPAE Recife', codigoUnidadeSaude: '260001', geres: 'I GERES', municipio: 'Recife', situacaoAtiva: true },
    { nomeDaUnidade: 'Hospital Regional do Agreste', codigoUnidadeSaude: '260002', geres: 'IV GERES', municipio: 'Caruaru', situacaoAtiva: true },
    { nomeDaUnidade: 'UPAE Petrolina', codigoUnidadeSaude: '260003', geres: 'VIII GERES', municipio: 'Petrolina', situacaoAtiva: true }
  ]);

  const [cardio, dermato, ultrassom] = await SchedulingItem.insertMany([
    { nomeDoItem: 'Consulta em Cardiologia', codigoOcupacaoProfissional: '225120', situacaoAtiva: true },
    { nomeDoItem: 'Consulta em Dermatologia', codigoOcupacaoProfissional: '225135', situacaoAtiva: true },
    { nomeDoItem: 'Ultrassonografia Abdome Total', codigoOcupacaoProfissional: '407010', situacaoAtiva: true }
  ]);

  const [profissional] = await Professional.insertMany([
    { nomeCompleto: 'Dra. Mariana Albuquerque', siglaConselho: 'CRM', numeroConselho: '12345-PE', unidadeVinculadaId: upaeRecife._id, codigoUnidadeSaude: upaeRecife.codigoUnidadeSaude, situacaoAtiva: true },
    { nomeCompleto: 'Dr. Rafael Nascimento', siglaConselho: 'CRM', numeroConselho: '67890-PE', unidadeVinculadaId: hospitalCaruaru._id, codigoUnidadeSaude: hospitalCaruaru.codigoUnidadeSaude, situacaoAtiva: true }
  ]);

  const senhaCriptografada = await bcrypt.hash('agendaorg123', 10);
  const users = await User.insertMany([
    { nomeCompleto: 'Admin AgendaOrg', login: 'admin', email: 'admin@agendaorg.local', senhaCriptografada, perfil: 'Administrador', situacaoAtiva: true },
    { nomeCompleto: 'Unidade Executante Recife', login: 'unidade', email: 'unidade@agendaorg.local', senhaCriptografada, perfil: 'Unidade Executante', unidadeId: upaeRecife._id, situacaoAtiva: true },
    { nomeCompleto: 'Apoiador da Regulação', login: 'apoiador', email: 'apoiador@agendaorg.local', senhaCriptografada, perfil: 'Apoiador da Regulação', situacaoAtiva: true },
    { nomeCompleto: 'Gestor da Regulação', login: 'gestor', email: 'gestor@agendaorg.local', senhaCriptografada, perfil: 'Gestor da Regulação', situacaoAtiva: true },
    { nomeCompleto: 'Gestor GERES', login: 'geres', email: 'geres@agendaorg.local', senhaCriptografada, perfil: 'Gestor/GERES', situacaoAtiva: true }
  ]);

  const agenda = await Agenda.create({
    unidadeId: upaeRecife._id,
    identificadorDaUnidade: 260001,
    mesCompetencia: '2026-06',
    estadoAtual: 'Recebida',
    observacoes: 'Agenda inicial para demonstração do fluxo SES-PE.',
    criadaPorId: users[1]._id,
    ofertas: [
      { data: '2026-06-18', horarioAtendimento: '08:00', quantidadeDeVagas: 12, itemAgendamentoId: cardio._id, profissionalId: profissional._id, indicativoAtiva: true },
      { data: '2026-06-19', horarioAtendimento: '10:00', quantidadeDeVagas: 8, itemAgendamentoId: dermato._id, profissionalId: profissional._id, indicativoAtiva: true },
      { data: '2026-06-20', horarioAtendimento: '13:00', quantidadeDeVagas: 6, itemAgendamentoId: ultrassom._id, profissionalId: profissional._id, indicativoAtiva: true }
    ]
  });

  await History.create({ agendaId: agenda._id, descricaoDaAcao: 'Agenda registrada e recebida para validação', idUsuarioResponsavel: users[1]._id, perfil: 'Unidade Executante' });
  await Notification.insertMany([{ usuarioId: users[2]._id, textoDaMensagem: 'Nova agenda recebida de UPAE Recife', agendaId: agenda._id }, { usuarioId: users[3]._id, textoDaMensagem: 'Painel gerencial atualizado com nova agenda', agendaId: agenda._id }]);
  await Parameter.insertMany([{ chaveDeIdentificacaoDaRegra: 'limiteArquivoMb', valorLimiteDefinido: 8, descricao: 'Tamanho maximo de anexo permitido em megabytes' }, { chaveDeIdentificacaoDaRegra: 'competenciaObrigatoria', valorLimiteDefinido: true, descricao: 'Toda agenda deve informar mes de competencia' }, { chaveDeIdentificacaoDaRegra: 'bloqueioExigeJustificativa', valorLimiteDefinido: true, descricao: 'Bloqueios de vagas devem conter justificativa auditavel' }]);
  await AuditLog.create({ usuarioId: users[0]._id, perfil: 'Administrador', descricaoDaAcao: 'Base inicial populada', recurso: 'seed' });

  console.log('Seed concluido. Logins: admin, unidade, apoiador, gestor, geres. Senha: agendaorg123');
  await mongoose.disconnect();
}

main().catch(async (error) => { console.error(error); await mongoose.disconnect(); process.exit(1); });

