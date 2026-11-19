package com.exemplo.etl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.sql.Types;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;

public class ImportadorPlanilha {

    private static final Properties PROPS = carregarProps();

    private static final String DB_URL = cfg("db.url",
            "jdbc:mysql://localhost:3306/signal_db?useSSL=false&allowPublicKeyRetrieval=true"
                    + "&serverTimezone=America/Sao_Paulo&characterEncoding=UTF-8&rewriteBatchedStatements=true");
    private static final String DB_USER = cfg("db.user", "root");
    private static final String DB_PASS = cfg("db.pass", "");
    private static final String SHEET_NAME = cfg("xlsx.sheet", "Netflix Engagement (plus)");

    private static final int BATCH_SIZE = 500;
    private static final int MAX_ERROS_NO_LOG = 100;
    private static final String ORIGEM_LOG = "ImportadorPlanilha";

    private static final String SQL_ABRIR_IMPORTACAO =
            "INSERT INTO importacao_planilha (fkUsuario, arquivo, dataa, status_importacao) "
                    + "VALUES (?, ?, ?, 'EM_ANDAMENTO')";

    private static final String SQL_FECHAR_IMPORTACAO =
            "UPDATE importacao_planilha SET dataFim = ?, linhasLidas = ?, linhasSucesso = ?, "
                    + "linhasErro = ?, status_importacao = ? WHERE idImportacao = ?";

    private static final String SQL_INSERT_DADO =
            "INSERT INTO dados_tratados (fkImportacao, linhaOrigem, titulo, disponivelGlobalmente, "
                    + "dataLancamento, horasAssistidas, numeroAvaliacoes, nota, generos, palavrasChave, descricao) "
                    + "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

    private static final String SQL_LOG =
            "INSERT INTO historico_log (nivel_log, mensagem, origem, data_hora, fkUsuario) "
                    + "VALUES (?, ?, ?, ?, ?)";

    private static final String H_TITLE = "title";
    private static final String H_GLOBAL = "available globally?";
    private static final String H_RELEASE = "release date";
    private static final String H_HOURS = "hours viewed";
    private static final String H_NUM_RATINGS = "number of ratings";
    private static final String H_RATING = "rating";
    private static final String H_GENRE = "genre";
    private static final String H_KEYWORDS = "key words";
    private static final String H_DESC = "description";

    private static final List<String> CABECALHOS_OBRIGATORIOS = List.of(
            H_TITLE, H_GLOBAL, H_RELEASE, H_HOURS, H_NUM_RATINGS, H_RATING, H_GENRE, H_KEYWORDS, H_DESC);

    static class Registro {
        int linha;
        String titulo;
        Boolean globalmente;
        LocalDate lancamento;
        Long horas;
        Integer numAvaliacoes;
        Integer nota;
        String generos;
        String palavrasChave;
        String descricao;
    }

    static class Resultado {
        int lidas;
        int sucesso;
        int erros;
        String erroFatal;

        void erro(int linha, String motivo) {
            erros++;
            if (erros <= MAX_ERROS_NO_LOG) {
                log("ERRO", "Linha " + linha + ": " + motivo);
            } else if (erros == MAX_ERROS_NO_LOG + 1) {
                log("AVISO", "Limite de " + MAX_ERROS_NO_LOG
                        + " erros detalhados atingido; os demais serão apenas contados.");
            }
        }
    }

    private static final DateTimeFormatter FMT_LOG = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    static class EntradaLog {
        final String nivel;
        final String mensagem;
        final LocalDateTime quando;

        EntradaLog(String nivel, String mensagem, LocalDateTime quando) {
            this.nivel = nivel;
            this.mensagem = mensagem;
            this.quando = quando;
        }
    }

    private static final List<EntradaLog> LOGS = new ArrayList<>();

    static void log(String nivel, String mensagem) {
        LocalDateTime agora = LocalDateTime.now();
        LOGS.add(new EntradaLog(nivel, mensagem, agora));
        String linha = String.format("[%s] [%-5s] %s", FMT_LOG.format(agora), nivel, mensagem);
        if ("ERRO".equals(nivel)) {
            System.err.println(linha);
        } else {
            System.out.println(linha);
        }
    }

    public static void main(String[] args) {
        Path arquivo = Path.of(args.length > 0 ? args[0] : cfg("xlsx.path", "dados/Book.xlsx"));
        int idUsuario = Integer.parseInt(cfg("import.usuario.id", "1"));
        LocalDateTime inicio = LocalDateTime.now();
        Resultado res = new Resultado();

        log("INFO", "Iniciando importação do arquivo '" + arquivo + "' (usuário " + idUsuario + ").");

        try (Connection con = DriverManager.getConnection(DB_URL, DB_USER, DB_PASS)) {
            con.setAutoCommit(false);
            log("INFO", "Conexão com o banco estabelecida.");

            boolean usuarioValido = false;
            Integer idImportacao = null;

            try {
                usuarioValido = usuarioExiste(con, idUsuario);
                if (!usuarioValido) {
                    throw new IllegalStateException("Usuário " + idUsuario
                            + " não existe na tabela usuario (ajuste import.usuario.id).");
                }
                if (!Files.isReadable(arquivo)) {
                    throw new IllegalStateException("Arquivo não encontrado ou ilegível: " + arquivo.toAbsolutePath());
                }

                idImportacao = abrirImportacao(con, idUsuario, arquivo, inicio);
                log("INFO", "Importação #" + idImportacao + " registrada em importacao_planilha.");

                List<Registro> validos = new ArrayList<>();
                extrair(arquivo, validos, res);
                carregar(con, idImportacao, validos, res);
            } catch (Exception e) {
                res.erroFatal = e.getClass().getSimpleName() + ": " + e.getMessage();
                log("ERRO", "Falha fatal: " + res.erroFatal);
                try {
                    con.rollback();
                } catch (SQLException ignored) {
                }
            }

            try {
                finalizar(con, idImportacao, usuarioValido ? idUsuario : null, arquivo, res);
            } catch (SQLException e) {
                log("ERRO", "Não foi possível gravar o log no banco: " + e.getMessage());
            }
        } catch (SQLException e) {
            log("ERRO", "Falha de conexão com o banco: " + e.getMessage());
            System.exit(1);
        }

        if (res.erroFatal != null) {
            System.exit(2);
        }
    }

    private static boolean usuarioExiste(Connection con, int idUsuario) throws SQLException {
        try (PreparedStatement ps = con.prepareStatement("SELECT 1 FROM usuario WHERE idUsuario = ?")) {
            ps.setInt(1, idUsuario);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }

    private static int abrirImportacao(Connection con, int idUsuario, Path arquivo, LocalDateTime inicio)
            throws SQLException {
        try (PreparedStatement ps = con.prepareStatement(SQL_ABRIR_IMPORTACAO, Statement.RETURN_GENERATED_KEYS)) {
            ps.setInt(1, idUsuario);
            ps.setString(2, arquivo.getFileName().toString());
            ps.setTimestamp(3, Timestamp.valueOf(inicio));
            ps.executeUpdate();
            try (ResultSet keys = ps.getGeneratedKeys()) {
                if (!keys.next()) {
                    throw new SQLException("Não foi possível obter o idImportacao gerado.");
                }
                int id = keys.getInt(1);
                con.commit();
                return id;
            }
        }
    }

    private static void finalizar(Connection con, Integer idImportacao, Integer fkUsuarioLog,
                                  Path arquivo, Resultado res) throws SQLException {
        LocalDateTime fim = LocalDateTime.now();
        String status = res.erroFatal != null ? "FALHA"
                : res.erros > 0 ? "PARCIAL" : "SUCESSO";

        if (idImportacao != null) {
            try (PreparedStatement ps = con.prepareStatement(SQL_FECHAR_IMPORTACAO)) {
                ps.setTimestamp(1, Timestamp.valueOf(fim));
                ps.setInt(2, res.lidas);
                ps.setInt(3, res.sucesso);
                ps.setInt(4, res.erros);
                ps.setString(5, status);
                ps.setInt(6, idImportacao);
                ps.executeUpdate();
            }
        }

        String resumo = String.format(
                "Importação #%s de '%s' finalizada com status %s: %d linha(s) lida(s), %d com sucesso, %d com erro.",
                idImportacao == null ? "-" : idImportacao.toString(),
                arquivo.getFileName(), status, res.lidas, res.sucesso, res.erros);
        log(res.erroFatal != null ? "ERRO" : (res.erros > 0 ? "AVISO" : "INFO"), resumo);

        try (PreparedStatement ps = con.prepareStatement(SQL_LOG)) {
            for (EntradaLog l : LOGS) {
                addLog(ps, l.nivel, l.mensagem, l.quando, fkUsuarioLog);
            }
            ps.executeBatch();
        }
        con.commit();
    }

    private static void addLog(PreparedStatement ps, String nivel, String mensagem,
                               LocalDateTime quando, Integer fkUsuario) throws SQLException {
        ps.setString(1, nivel);
        ps.setString(2, mensagem);
        ps.setString(3, ORIGEM_LOG);
        ps.setTimestamp(4, Timestamp.valueOf(quando));
        if (fkUsuario == null) ps.setNull(5, Types.INTEGER); else ps.setInt(5, fkUsuario);
        ps.addBatch();
    }

    static void extrair(Path arquivo, List<Registro> destino, Resultado res) throws Exception {

        try (Workbook wb = WorkbookFactory.create(arquivo.toFile(), null, true)) {
            Sheet sheet = wb.getSheet(SHEET_NAME);
            if (sheet == null) {
                List<String> nomes = new ArrayList<>();
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    nomes.add(wb.getSheetName(i));
                }
                throw new IllegalStateException("Aba '" + SHEET_NAME + "' não encontrada. Abas existentes: " + nomes);
            }

            Row cabecalho = sheet.getRow(sheet.getFirstRowNum());
            if (cabecalho == null) {
                throw new IllegalStateException("A aba '" + SHEET_NAME + "' está vazia.");
            }

            Map<String, Integer> idx = new HashMap<>();
            for (Cell c : cabecalho) {
                String nome = texto(c);
                if (nome != null) {
                    idx.put(nome.trim().toLowerCase(Locale.ROOT), c.getColumnIndex());
                }
            }
            for (String h : CABECALHOS_OBRIGATORIOS) {
                if (!idx.containsKey(h)) {
                    throw new IllegalStateException("Coluna obrigatória ausente na planilha: '" + h + "'");
                }
            }

            log("INFO", "Lendo aba '" + SHEET_NAME + "' (colunas validadas, "
                    + (sheet.getLastRowNum() - sheet.getFirstRowNum()) + " linha(s) abaixo do cabeçalho).");

            for (int i = sheet.getFirstRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || linhaVazia(row)) {
                    continue;
                }
                res.lidas++;
                try {
                    destino.add(mapear(row, i + 1, idx));
                } catch (RuntimeException e) {
                    res.erro(i + 1, e.getMessage());
                }
            }
            log("INFO", String.format("Leitura concluída: %d linha(s) lida(s), %d válida(s), %d rejeitada(s).",
                    res.lidas, destino.size(), res.erros));
        }
    }

    private static Registro mapear(Row row, int numeroLinha, Map<String, Integer> idx) {
        Registro r = new Registro();
        r.linha = numeroLinha;

        r.titulo = textoLimpo(cel(row, idx, H_TITLE));
        if (r.titulo == null) {
            throw new IllegalArgumentException("Title vazio");
        }
        if (r.titulo.length() > 300) {
            throw new IllegalArgumentException("Title excede 300 caracteres");
        }

        r.globalmente = booleano(cel(row, idx, H_GLOBAL));
        r.lancamento = data(cel(row, idx, H_RELEASE));

        r.horas = inteiroLongo(cel(row, idx, H_HOURS));
        if (r.horas == null) {
            throw new IllegalArgumentException("Hours Viewed ausente");
        }

        Long numAval = inteiroLongo(cel(row, idx, H_NUM_RATINGS));
        if (numAval != null) {
            r.numAvaliacoes = Math.toIntExact(numAval);
        }

        Long nota = inteiroLongo(cel(row, idx, H_RATING));
        if (nota != null) {
            if (nota < 0 || nota > 100) {
                throw new IllegalArgumentException("Rating fora do intervalo 0-100: " + nota);
            }
            r.nota = nota.intValue();
        }

        String generos = textoLimpo(cel(row, idx, H_GENRE));
        if (generos != null) {
            generos = generos.replaceAll("[\\[\\]']", "").trim();
            r.generos = generos.isEmpty() ? null : generos;
        }
        if (r.generos != null && r.generos.length() > 100) {
            throw new IllegalArgumentException("Genre excede 100 caracteres");
        }

        r.palavrasChave = textoLimpo(cel(row, idx, H_KEYWORDS));
        r.descricao = textoLimpo(cel(row, idx, H_DESC));
        return r;
    }

    static void carregar(Connection con, int idImportacao, List<Registro> registros, Resultado res)
            throws SQLException {
        int totalLotes = (registros.size() + BATCH_SIZE - 1) / BATCH_SIZE;
        log("INFO", "Inserindo " + registros.size() + " registro(s) em dados_tratados ("
                + totalLotes + " lote(s) de até " + BATCH_SIZE + ").");

        try (PreparedStatement ps = con.prepareStatement(SQL_INSERT_DADO)) {
            for (int i = 0; i < registros.size(); i += BATCH_SIZE) {
                List<Registro> lote = registros.subList(i, Math.min(i + BATCH_SIZE, registros.size()));
                int numLote = i / BATCH_SIZE + 1;
                try {
                    for (Registro r : lote) {
                        preencher(ps, idImportacao, r);
                        ps.addBatch();
                    }
                    ps.executeBatch();
                    con.commit();
                    res.sucesso += lote.size();
                    log("INFO", String.format("Lote %d/%d gravado (%d/%d registros).",
                            numLote, totalLotes, res.sucesso, registros.size()));
                } catch (SQLException e) {

                    log("AVISO", "Lote " + numLote + "/" + totalLotes + " falhou (" + e.getMessage()
                            + "); reprocessando linha a linha.");
                    con.rollback();
                    ps.clearBatch();
                    for (Registro r : lote) {
                        try {
                            preencher(ps, idImportacao, r);
                            ps.executeUpdate();
                            con.commit();
                            res.sucesso++;
                        } catch (SQLException e2) {
                            con.rollback();
                            res.erro(r.linha, "SQL: " + e2.getMessage());
                        }
                    }
                }
            }
        }
    }

    private static void preencher(PreparedStatement ps, int idImportacao, Registro r) throws SQLException {
        ps.setInt(1, idImportacao);
        ps.setInt(2, r.linha);
        ps.setString(3, r.titulo);
        if (r.globalmente == null) ps.setNull(4, Types.BOOLEAN); else ps.setBoolean(4, r.globalmente);
        if (r.lancamento == null) ps.setNull(5, Types.DATE); else ps.setDate(5, Date.valueOf(r.lancamento));
        ps.setLong(6, r.horas);
        if (r.numAvaliacoes == null) ps.setNull(7, Types.INTEGER); else ps.setInt(7, r.numAvaliacoes);
        if (r.nota == null) ps.setNull(8, Types.INTEGER); else ps.setInt(8, r.nota);
        ps.setString(9, r.generos);
        ps.setString(10, r.palavrasChave);
        ps.setString(11, r.descricao);
    }

    private static Cell cel(Row row, Map<String, Integer> idx, String cabecalho) {
        return row.getCell(idx.get(cabecalho), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
    }

    private static boolean linhaVazia(Row row) {
        for (Cell c : row) {
            if (textoLimpo(c) != null) {
                return false;
            }
        }
        return true;
    }

    private static CellType tipo(Cell c) {
        return c.getCellType() == CellType.FORMULA ? c.getCachedFormulaResultType() : c.getCellType();
    }

    static String texto(Cell c) {
        if (c == null) {
            return null;
        }
        switch (tipo(c)) {
            case STRING:
                return c.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c)) {
                    return c.getLocalDateTimeCellValue().toString();
                }

                return BigDecimal.valueOf(c.getNumericCellValue()).stripTrailingZeros().toPlainString();
            case BOOLEAN:
                return String.valueOf(c.getBooleanCellValue());
            default:
                return null;
        }
    }

    static String textoLimpo(Cell c) {
        String s = texto(c);
        if (s == null) {
            return null;
        }
        s = s.replace('\u00A0', ' ').trim();
        s = s.replace("&apos;", "'").replace("&#39;", "'").replace("&quot;", "\"")
                .replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&");
        return s.isEmpty() ? null : s;
    }

    static String somenteDigitos(String s) {
        if (s == null) {
            return null;
        }
        String d = s.replaceAll("\\D", "");
        return d.isEmpty() ? null : d;
    }

    static Long inteiroLongo(Cell c) {
        String s = textoLimpo(c);
        if (s == null) {
            return null;
        }
        try {
            BigDecimal bd = new BigDecimal(s.replace(",", "."));
            return bd.setScale(0, RoundingMode.HALF_UP).longValueExact();
        } catch (NumberFormatException | ArithmeticException e) {
            throw new IllegalArgumentException("Valor numérico inválido: '" + s + "'");
        }
    }

    static Boolean booleano(Cell c) {
        String s = textoLimpo(c);
        if (s == null) {
            return null;
        }
        switch (s.toLowerCase(Locale.ROOT)) {
            case "yes": case "sim": case "true": case "1": case "y": case "s":
                return Boolean.TRUE;
            case "no": case "nao": case "não": case "false": case "0": case "n":
                return Boolean.FALSE;
            default:
                throw new IllegalArgumentException("Valor booleano inválido: '" + s + "'");
        }
    }

    private static final DateTimeFormatter[] FORMATOS_DATA = {
            DateTimeFormatter.ofPattern("dd/MM/yyyy"),
            DateTimeFormatter.ISO_LOCAL_DATE,
            DateTimeFormatter.ofPattern("MM/dd/yyyy")
    };

    static LocalDate data(Cell c) {
        if (c == null) {
            return null;
        }
        if (tipo(c) == CellType.NUMERIC) {
            double v = c.getNumericCellValue();
            if (DateUtil.isCellDateFormatted(c) || DateUtil.isValidExcelDate(v)) {
                return DateUtil.getJavaDate(v).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            }
            throw new IllegalArgumentException("Data inválida: " + v);
        }
        String s = textoLimpo(c);
        if (s == null) {
            return null;
        }
        if (s.length() > 10) {
            s = s.substring(0, 10);
        }
        for (DateTimeFormatter f : FORMATOS_DATA) {
            try {
                return LocalDate.parse(s, f);
            } catch (DateTimeParseException ignored) {
            }
        }
        throw new IllegalArgumentException("Data inválida: '" + s + "'");
    }

    private static Properties carregarProps() {
        Properties p = new Properties();
        Path f = Path.of("config.properties");
        if (Files.isReadable(f)) {
            try (InputStream in = Files.newInputStream(f)) {
                p.load(in);
            } catch (IOException e) {
                System.err.println("Aviso: não foi possível ler config.properties: " + e.getMessage());
            }
        }
        return p;
    }

    private static String cfg(String chave, String padrao) {
        String env = System.getenv(chave.toUpperCase(Locale.ROOT).replace('.', '_'));
        if (env != null && !env.isBlank()) {
            return env;
        }
        String v = PROPS.getProperty(chave);
        return (v == null || v.isBlank()) ? padrao : v.trim();
    }
}