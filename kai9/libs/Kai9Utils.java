package kai9.libs;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringReader;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.function.Consumer;

import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.spec.SecretKeySpec;
import javax.servlet.http.HttpServletResponse;

import org.jasypt.encryption.pbe.StandardPBEStringEncryptor;
import org.jasypt.iv.RandomIvGenerator;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.config.YamlPropertiesFactoryBean;
import org.springframework.core.NestedRuntimeException;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PropertiesLoaderUtils;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Component;
import org.springframework.transaction.TransactionDefinition;
import org.springframework.transaction.TransactionStatus;
import org.springframework.transaction.support.DefaultTransactionDefinition;
import org.springframework.transaction.support.TransactionTemplate;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * ユーティリティ
 */
@Component
public class Kai9Utils {

    // Kai9ComのDB制御をDI
    private static JdbcTemplate jdbcTemplate_com;

    // Kai9ComのDB設定をapplication.propertyesからロード
    @Autowired
    public void setJdbcTemplateCom(@Qualifier("commonjdbc")
    JdbcTemplate jdbcTemplate_com) {
        Kai9Utils.jdbcTemplate_com = jdbcTemplate_com;
    }

    /**
     * 例外を処理し、JsonResponseを生成するメソッド
     * 
     * @param e 発生した例外
     * @param res HttpServletResponseオブジェクト
     * @return 生成されたJsonResponseオブジェクト
     * @throws IOException 入出力エラーが発生した場合
     * @throws JSONException JSONエラーが発生した場合
     */
    public static JsonResponse handleException(Exception e, HttpServletResponse res) throws IOException, JSONException {
        JsonResponse json = new JsonResponse();
        String msg = GetException(e);
        json.setReturn_code(HttpStatus.INTERNAL_SERVER_ERROR.value());
        json.setMsg(msg);
        json.SetJsonResponse(res);
        return json;
    }

    /**
     * 例外を処理し、JsonResponseをResourceに変換してResponseEntityとして返すメソッド。
     * handleExceptionメソッドを使用してJsonResponseを生成し、それをResourceに変換します。
     * 
     * @param e 発生した例外
     * @param res HttpServletResponseオブジェクト
     * @param status レスポンスのHttpStatus
     * @return Resourceを含むResponseEntity
     * @throws IOException 例外処理中に入出力エラーが発生した場合
     */
    public static ResponseEntity<Resource> handleExceptionAsResponseEntity(Exception e, HttpServletResponse res, HttpStatus status) throws IOException {
        // 例外メッセージを取得
        String Msg = Kai9Utils.GetException(e);

        // ダミーのJSONデータとしてメッセージを埋め込む
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("msg", Msg);

        // JSON文字列に変換
        String jsonString = jsonObject.toString();

        // JSON文字列をバイト配列に変換し、Blobとして扱う
        ByteArrayResource resource = new ByteArrayResource(jsonString.getBytes(StandardCharsets.UTF_8));

        // Content-DispositionヘッダーでBlobとして返すことを明示
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=errorMessage.json");
        headers.add(HttpHeaders.CONTENT_TYPE, "application/json");

        // Blobを含むResponseEntityを返す
        return new ResponseEntity<>(resource, headers, status);
    }

    /**
     * 例外メッセージの出方を、ユーザの権限で変える。管理者権限の場合、詳細なスタックトレースを表示する
     */
    public static String GetException(Exception e) {
        String crlf = System.lineSeparator(); // 改行コード

        // 認証ユーザ取得
        int authority_lv = 2; // 一番権限の低い参照専用をデフォルトにする
        try {
            Authentication auth = SecurityContextHolder.getContext().getAuthentication();
            String name = auth.getName();
            authority_lv = getauthority_lvByLoginID(name);
        } catch (Exception exc) {
            // anonymousUserの場合、アベンドするので、例外を無視してauthority_lv = 2を維持する
        }

        // 出力するための StringBuilder を準備
        StringBuilder message = new StringBuilder();

        // 例外の原因をリストに格納
        List<Throwable> causes = new ArrayList<>();
        Throwable cause = e;
        while (cause != null) {
            causes.add(cause);
            cause = cause.getCause();
        }

        // 【エラーが発生しました】部分
        message.append("【エラーが発生しました】").append(crlf);

        // 直接原因部分（リストの最後の要素から順に出力）
        for (int i = causes.size() - 1; i >= 0; i--) {
            Throwable currentCause = causes.get(i);
            message.append("エラー").append(causes.size() - i).append("：")
                    .append(currentCause.getClass().getName()).append(": ")
                    .append(currentCause.getMessage()).append(crlf + crlf);
        }

        // 管理者権限がある場合、詳細部分を追加
        if (authority_lv == 3) {
            // 詳細部分もリストの最後の要素から順に出力
            int detailCount = 1;
            for (int i = causes.size() - 1; i >= 0; i--) {
                Throwable currentCause = causes.get(i);
                StringWriter sw = new StringWriter();
                PrintWriter pw = new PrintWriter(sw);
                currentCause.printStackTrace(pw);
                message.append("【詳細").append(detailCount).append("】").append(crlf);

                // スタックトレース全体から詳細を取得（2行目以降）
                try (BufferedReader br = new BufferedReader(new StringReader(sw.toString()))) {
                    br.readLine(); // 1行目はスキップ（既に取得済み）
                    String line;
                    while ((line = br.readLine()) != null) {
                        message.append(line).append(crlf);
                    }
                    message.append(crlf);
                } catch (IOException ioException) {
                    // 例外処理（通常はここには来ないはず）
                }

                detailCount++;
            }
        }

        return message.toString();
    }

    /**
     * 例外メッセージ中の特殊文字をエスケープ
     * 例外文字に制御文字が入るとReact側で制御不能に陥るのでエスケープする
     */
    public static String escapeMessage(String message) {
        if (message == null) {
            return null;
        }
        return message.replaceAll("%", "%25")
                .replaceAll("\n", "%0A")
                .replaceAll("\r", "%0D")
                .replaceAll("\t", "%09")
                .replaceAll("\b", "%08")
                .replaceAll("\f", "%0C")
                .replaceAll("\\\\", "%5C")
                .replaceAll("\"", "%22")
                .replaceAll("'", "%27")
                .replaceAll("#", "%23")
                .replaceAll("&", "%26")
                .replaceAll("<", "%3C")
                .replaceAll(">", "%3E");
    }

    /**
     * ログインIDからユーザの権限を取得するメソッド
     * 
     * @param login_id ログインID
     * @return int ユーザの権限レベル(1:一般、2:参照専用、3:管理者)
     */
    public static int getauthority_lvByLoginID(String login_id) {
        String sql = "select * from m_user_a where login_id = ?";
        // JdbcTemplateを利用して、SQLを実行。取得したMapから権限レベルを返す
        Map<String, Object> map = jdbcTemplate_com.queryForMap(sql, login_id);
        int authority_lv = (int) map.get("authority_lv");
        return authority_lv;
    }

    /**
     * 引数のオブジェクトをJSON文字列に変換
     */
    public static String getJsonData(Object data) {
        String retVal = null;
        ObjectMapper objectMapper = new ObjectMapper();
        try {
            retVal = objectMapper.writeValueAsString(data);
        } catch (JsonProcessingException e) {
            System.err.println(e);
        }
        return retVal;
    }

    /**
     * サイズの大きなカラム(BLOBや、長い文字列)がある場合、そのカラムをNULLで取得するSQLを動的に作成
     * 
     * テーブル名、カラム名、テーブル接頭子を受け取り、
     * 指定されたテーブルからカラム名を取得して、SELECT文を作成する。
     * ただし、columnNameToBeNullで指定されたカラム名にはNULLを返す。
     * カラム名にスペースが含まれている場合は、
     * テーブル接頭子を付加してからカラム名を返す。
     *
     * @param tableName SELECT対象のテーブル名
     * @param columnNameToBeNull NULL値にしたいカラム名
     * @param tablePrefix テーブル接頭子（null可）
     * @param ReplaceText 代わりに入れたい文字(nullや、任意の文字等)
     * @param jdbcTemplate_param JdbcTemplate(接続先DBに合わせれる様に)
     * @return SELECT文
     * @throws SQLException データベースアクセスエラー
     */
    public static String createSelectSqlWithNullColumn(String tableName, String columnNameToBeNull, String tablePrefix, String ReplaceText, JdbcTemplate jdbcTemplate_param) throws SQLException {
        // 指定されたテーブルからカラム名を取得する
        List<String> columnNames = getColumnNamesWithoutSystemTables(tableName, jdbcTemplate_param);

        StringBuilder selectSql = new StringBuilder();
        for (int i = 0; i < columnNames.size(); i++) {
            // 最初のカラム名以外は、","で区切る
            if (i > 0) {
                selectSql.append(", ");
            }
            // columnNameToBeNullで指定されたカラム名の場合は、"NULL"等の、指定されたReplaceTextを追加する
            if (columnNames.get(i).equalsIgnoreCase(columnNameToBeNull)) {
                selectSql.append(ReplaceText + " AS ");
            } else {
                // tablePrefixが指定されている場合は、カラム名の先頭に接頭子を追加する
                if (tablePrefix != null) {
                    selectSql.append(tablePrefix).append(".");
                }
            }
            selectSql.append(columnNames.get(i));
        }

        return selectSql.toString();
    }

    /**
     * ↑のcreateSelectSqlWithNullColumnで利用
     * 
     * テーブルからカラム名を取得する。
     * カラム名にスペースが含まれている場合は、
     * テーブル接頭子を付加してからカラム名を返す。
     *
     * @param tableName カラム名を取得するテーブル名
     * @return カラム名のリスト
     * @throws SQLException データベースアクセスエラー
     */
    private static List<String> getColumnNamesWithoutSystemTables(String tableName, JdbcTemplate jdbcTemplate_param) throws SQLException {
        List<String> columnNames = new ArrayList<>();
        Connection connection = jdbcTemplate_param.getDataSource().getConnection();
        // LIMIT 0を指定することで、SQL文を実行しないで結果セットのメタデータを取得できる
        try (ResultSet rs = connection.createStatement().executeQuery("SELECT * FROM " + tableName + " LIMIT 0")) {
            ResultSetMetaData metaData = rs.getMetaData();
            int columnCount = metaData.getColumnCount();
            for (int i = 1; i <= columnCount; i++) {
                String columnName = metaData.getColumnName(i);
                // カラム名にスペースが含まれている場合は、接頭子を付加してからカラム名を返す
                if (columnName.contains(" ")) {
                    String[] parts = columnName.split(" ");
                    columnName = parts[0] + "." + parts[1];
                }
                columnNames.add(columnName);
            }
        }
        return columnNames;
    }

    /**
     * トラザクション制御
     */

    // トランザクションの開始
    public static TransactionStatus beginTransaction(TransactionTemplate transactionTemplate) {
        // トランザクションの定義を作成する
        DefaultTransactionDefinition def = new DefaultTransactionDefinition();
        // トランザクションの伝搬動作を「常に新しいトランザクションを開始する」に設定する
        def.setPropagationBehavior(TransactionDefinition.PROPAGATION_REQUIRES_NEW);

        // トランザクションマネージャーからトランザクションを取得して返す
        return transactionTemplate.getTransactionManager().getTransaction(def);
    }

    // トランザクションのコミット
    public static void commitTransaction(TransactionTemplate transactionTemplate, TransactionStatus status) {
        transactionTemplate.getTransactionManager().commit(status);
    }

    // トランザクションのロールバック
    public static void rollbackTransaction(TransactionTemplate transactionTemplate, TransactionStatus status) {
        transactionTemplate.getTransactionManager().rollback(status);
    }

    /**
     * loggerでのログ出力
     *
     * @param className ログを出力するクラス名
     * @param msg ログメッセージ
     * @param type ログの種類（"info"、"warn"、"error"のいずれか）
     */
    public static void makeLog(String type, String msg, Class<?> clazz) {
        String MSG = "【Kai9】" + msg;
        Logger logger = LoggerFactory.getLogger(clazz);
        if (type.equals("info")) logger.info(MSG);
        if (type.equals("warn")) logger.warn(MSG);
        if (type.equals("error")) logger.error(MSG);
    }
    
    /**
     * 例外の原因を収集し、メッセージとスタックトレースを一括で生成するメソッド
     */
    public static String processExceptionMessages(Throwable e) {
        String crlf = System.lineSeparator(); // 改行コード
        StringBuilder message = new StringBuilder();
        List<Throwable> causes = new ArrayList<>();
        Throwable cause = e;

        while (cause != null) {
            causes.add(cause);
            cause = cause.getCause();
        }

        // 直接原因部分（リストの最後の要素から順に出力）
        for (int i = causes.size() - 1; i >= 0; i--) {
            Throwable currentCause = causes.get(i);
            message.append("エラー").append(causes.size() - i).append("：")
                   .append(currentCause.getClass().getName()).append(": ")
                   .append(currentCause.getMessage()).append(crlf).append(crlf);
        }

        // 詳細部分を追加
        int detailCount = 1;
        for (int i = causes.size() - 1; i >= 0; i--) {
            Throwable currentCause = causes.get(i);
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            currentCause.printStackTrace(pw);
            message.append("【詳細").append(detailCount).append("】").append(crlf);

            // スタックトレース全体から詳細を取得（2行目以降）
            try (BufferedReader br = new BufferedReader(new StringReader(sw.toString()))) {
                br.readLine(); // 1行目はスキップ（既に取得済み）
                String line;
                while ((line = br.readLine()) != null) {
                    message.append(line).append(crlf);
                }
                message.append(crlf);
            } catch (IOException ioException) {
                // 例外処理（通常はここには来ないはず）
            }

            detailCount++;
        }

        return message.toString();
    }
    

    /**
     * application.propertiesから指定されたプロパティを取得する
     *
     * @param propertyName 取得したいプロパティの名前
     * @return プロパティの値
     */
    public static String getPropertyFromProperties(String propertyName) {
        try {
            Resource resource = new ClassPathResource("application.properties");
            Properties properties = PropertiesLoaderUtils.loadProperties(resource);
            return properties.getProperty(propertyName);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * application.ymlから指定されたプロパティを取得する
     *
     * @param propertyName 取得したいプロパティの名前
     * @return プロパティの値
     */
    public static String getPropertyFromYaml(String propertyName) {
        YamlPropertiesFactoryBean factory = new YamlPropertiesFactoryBean();
        factory.setResources(new ClassPathResource("application.yml"));
        Properties properties = factory.getObject();

        String propertyValue = properties.getProperty(propertyName);

        // プロパティが暗号化されている場合に複合化を行う
        if (propertyValue != null && propertyValue.startsWith("ENC(") && propertyValue.endsWith(")")) {
            propertyValue = decryptPropertyValue(propertyValue);
        }

        return propertyValue; // 複合化された値またはそのままの値を返す
    }

    public static String decryptPropertyValue(String encryptedValue) {
        // JasyptのEncryptorの設定。パスワードは環境変数やシステムプロパティから取得
        StandardPBEStringEncryptor encryptor = new StandardPBEStringEncryptor();
        encryptor.setAlgorithm("PBEWITHHMACSHA512ANDAES_256");
        encryptor.setIvGenerator(new RandomIvGenerator()); // IVジェネレータ

        // 環境変数からパスワードを取得。設定されていない場合はエラーをスロー
        String password = System.getenv("JASYPT_ENCRYPTOR_PASSWORD");
        if (password == null) {
            throw new IllegalStateException("環境変数 'JASYPT_ENCRYPTOR_PASSWORD' が設定されていません。");
        }
        encryptor.setPassword(password);

        // encryptedValueがnullでないか、"ENC("と")"で囲まれているかを確認
        if (encryptedValue == null || !encryptedValue.startsWith("ENC(") || !encryptedValue.endsWith(")")) {
            throw new IllegalArgumentException("暗号化された値の形式が無効です。");
        }

        try {
            // "ENC("と")"を取り除いて複合化する
            String decryptedValue = encryptor.decrypt(encryptedValue.substring(4, encryptedValue.length() - 1));
            if (decryptedValue == null) {
                throw new IllegalStateException("複合化に失敗しました。パスワードまたは暗号化データを確認してください。");
            }
            return decryptedValue;
        } catch (Exception e) {
            // 詳細なエラーメッセージを出力して問題を特定
            throw new IllegalStateException("複合化処理中にエラーが発生しました: " + e.getMessage(), e);
        }
    }


    // 暗号化(可逆性)
    public static void encryptAndSet(Consumer<String> pwSetter, Consumer<String> keySetter, String plainText) throws Exception {
        // シークレットキー作成
        KeyGenerator keyGen = KeyGenerator.getInstance("AES");
        keyGen.init(256);
        SecretKey secretKey = keyGen.generateKey();

        // 暗号化 (ECBモードを使用)
        Cipher cipher = Cipher.getInstance("AES/ECB/PKCS5Padding");
        cipher.init(Cipher.ENCRYPT_MODE, secretKey);
        byte[] encryptedBytes = cipher.doFinal(plainText.getBytes());
        String encryptedText = Base64.getEncoder().encodeToString(encryptedBytes);

        // SecretKeyをBase64でエンコードしてStringとして保存
        String encodedKey = Base64.getEncoder().encodeToString(secretKey.getEncoded());

        // セッターを呼び出して値をセット
        pwSetter.accept(encryptedText);
        keySetter.accept(encodedKey);
    }

    // 複合化
    public static String decrypt(String encryptedText, String encodedKey) throws Exception {
        // SecretKeyをBase64からデコードして復元
        byte[] decodedKey = Base64.getDecoder().decode(encodedKey);
        SecretKeySpec secretKey = new SecretKeySpec(decodedKey, 0, decodedKey.length, "AES");

        // 暗号化されたテキストを復号化 (ECBモードを使用)
        Cipher cipher = Cipher.getInstance("AES/ECB/PKCS5Padding");
        cipher.init(Cipher.DECRYPT_MODE, secretKey);
        byte[] decryptedBytes = cipher.doFinal(Base64.getDecoder().decode(encryptedText));
        return new String(decryptedBytes);
    }
    
    
    // ヒープメモリの情報をログに出力
    public static void logMemoryInfo(Class<?> clazz) {
        // Runtime インスタンスを取得
        Runtime runtime = Runtime.getRuntime();

        // ヒープの最大サイズ (bytes)
        long maxMemory = runtime.maxMemory();
        // ヒープの初期サイズ (bytes)
        long totalMemory = runtime.totalMemory();
        // 現在の使用中メモリ (bytes)
        long usedMemory = totalMemory - runtime.freeMemory();

        // ログ出力 (最大ヒープサイズ, 初期ヒープサイズ, 使用中メモリ)
        Kai9Utils.makeLog("info", "最大ヒープサイズ (Xmx): " + (maxMemory / 1024 / 1024) + " MB", clazz);
        Kai9Utils.makeLog("info", "現在のヒープサイズ (Xms): " + (totalMemory / 1024 / 1024) + " MB", clazz);
        Kai9Utils.makeLog("info", "使用中メモリ: " + (usedMemory / 1024 / 1024) + " MB", clazz);
    }
    
    // ローカルコマンドを実行し、結果を取得するメソッド
    public static String executeLocalCommand(String command) {
        StringBuilder output = new StringBuilder();
        try {
            // コマンドを実行する
            ProcessBuilder builder = new ProcessBuilder("powershell.exe", command);
            builder.redirectErrorStream(true);
            Process process = builder.start();

            // 結果を取得する
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append("\n");
            }

            // プロセス終了を待つ
            process.waitFor();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return output.toString().trim(); // 出力を返す
    }
    
    
}


