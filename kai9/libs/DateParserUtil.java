package kai9.libs;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAccessor;
import java.time.temporal.TemporalQueries;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

/**
 * DateParserUtil クラスは、さまざまなフォーマットの日付文字列を Date オブジェクトに変換するための共通関数
 */
public class DateParserUtil {

    // 使用する日付フォーマットのリスト
    private static final List<DateTimeFormatter> DATE_FORMATTERS = new ArrayList<>();
    private static final ConcurrentMap<Integer, DateTimeFormatter> SUCCESSFUL_FORMATTERS = new ConcurrentHashMap<>();

    // 静的イニシャライザで日付フォーマットをリストに追加
    // react側のvalidateExcelFile.isValidDateと平仄を合わせる事(同じフォーマッタを定義しているので
    static {
        // ISO形式
        DATE_FORMATTERS.add(DateTimeFormatter.ISO_DATE_TIME);
        // 年月日と時分秒を含む形式
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSSX"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSS"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ssX"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd'T'HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.S"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SS"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SSS"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-d HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-M-dd HH:mm:ss"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-M-d HH:mm:ss"));
        // 年月日形式
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("MM/dd/yyyy"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyyMMdd"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("dd-MM-yyyy"));
        // 英語の曜日を含む形式
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("EEE MMM dd HH:mm:ss zzz yyyy", Locale.ENGLISH));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("EEE MMM dd yyyy", Locale.ENGLISH));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("EEE, dd MMM yyyy HH:mm:ss z", Locale.ENGLISH));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("EEEE, dd MMMM yyyy HH:mm:ss z", Locale.ENGLISH));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("EEE, dd MMM yyyy HH:mm:ss z", Locale.ENGLISH));
        // その他の形式
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("d MMM uuuu"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("d MMMM uuuu"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("MMM d, uuuu"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/d"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/M/dd"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/M/d"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-M-d"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy-M-dd"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd'T'HH:mm"));
        DATE_FORMATTERS.add(DateTimeFormatter.ofPattern("yyyy/MM/dd'T'HH:mm:ss"));

    }

    /**
     * 入力文字列をさまざまな日付フォーマットで解析し、Date オブジェクトに変換
     *
     * @param dateStr 解析する日付文字列
     * @return 解析された Date オブジェクト
     * @throws IllegalArgumentException 日付文字列がどのフォーマットでも解析できない場合
     */
    public static Date strToDate(String dateStr) {
        int dateStrHash = dateStr.hashCode();
        // 成功したフォーマットが記憶されている場合、それを優先的に使用する
        if (SUCCESSFUL_FORMATTERS.containsKey(dateStrHash)) {
            try {
                DateTimeFormatter cachedFormatter = SUCCESSFUL_FORMATTERS.get(dateStrHash);
                TemporalAccessor temporalAccessor = cachedFormatter.parse(dateStr);
                return convertTemporalAccessorToDate(temporalAccessor);
            } catch (DateTimeParseException e) {
                // キャッシュされたフォーマットが失敗した場合、通常の処理に戻る
            }
        }

        for (DateTimeFormatter formatter : DATE_FORMATTERS) {
            try {
                TemporalAccessor temporalAccessor = formatter.parseBest(dateStr, ZonedDateTime::from, LocalDateTime::from, LocalDate::from);
                Date parsedDate = convertTemporalAccessorToDate(temporalAccessor);
                SUCCESSFUL_FORMATTERS.put(dateStrHash, formatter); // 成功したフォーマットを記憶
                return parsedDate;
            } catch (DateTimeParseException e) {
                // 無視して次のフォーマットを試す
            }
        }
        // どのフォーマットでも解析できなかった場合は例外を投げる
        throw new IllegalArgumentException("日付から文字列への変換に失敗しました。DateParserUtilの日付フォーマットリストにパターンを追記すれば変換できるようになります。: " + dateStr);
    }

    /**
     * TemporalAccessor インスタンスを Date オブジェクトに変換
     *
     * このメソッドは、TemporalAccessor インスタンスを受け取り、それを java.util.Date オブジェクトに
     * 変換します。ZonedDateTime、LocalDateTime、および LocalDate のインスタンスを処理し、
     * 最初に型が特定できない場合には TemporalAccessor から特定の時制情報を問い合わせる
     *
     * @param temporalAccessor 変換する TemporalAccessor インスタンス
     * @return 変換された Date オブジェクト
     * @throws IllegalArgumentException 未知の TemporalAccessor タイプの場合にスローされる
     */
    private static Date convertTemporalAccessorToDate(TemporalAccessor temporalAccessor) {
        // temporalAccessor が ZonedDateTime の場合
        if (temporalAccessor instanceof ZonedDateTime) {
            return Date.from(((ZonedDateTime) temporalAccessor).toInstant());
            // temporalAccessor が LocalDateTime の場合
        } else if (temporalAccessor instanceof LocalDateTime) {
            return Date.from(((LocalDateTime) temporalAccessor).atZone(ZoneId.systemDefault()).toInstant());
            // temporalAccessor が LocalDate の場合
        } else if (temporalAccessor instanceof LocalDate) {
            return Date.from(((LocalDate) temporalAccessor).atStartOfDay(ZoneId.systemDefault()).toInstant());
            // temporalAccessor にタイムゾーン情報が含まれる場合
        } else if (temporalAccessor.query(TemporalQueries.zone()) != null) {
            return Date.from(ZonedDateTime.from(temporalAccessor).toInstant());
            // temporalAccessor にローカル時間情報が含まれる場合
        } else if (temporalAccessor.query(TemporalQueries.localTime()) != null) {
            return Date.from(LocalDateTime.from(temporalAccessor).atZone(ZoneId.systemDefault()).toInstant());
//        	LocalDateTime localDateTime = LocalDateTime.of(
//        		    temporalAccessor.query(TemporalQueries.localDate()), 
//        		    temporalAccessor.query(TemporalQueries.localTime())
//        		);
//        	return Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
            // temporalAccessor にローカル日付情報が含まれる場合
        } else if (temporalAccessor.query(TemporalQueries.localDate()) != null) {
            return Date.from(LocalDate.from(temporalAccessor).atStartOfDay(ZoneId.systemDefault()).toInstant());
        }
        // 未知の TemporalAccessor タイプの場合は例外をスロー
        throw new IllegalArgumentException("不明な TemporalAccessor 型です");
    }

    /**
     * Date オブジェクトを指定されたフォーマットの文字列に変換
     *
     * @param date 変換する Date オブジェクト
     * @param format 変換するフォーマット
     * @return 変換された日付文字列
     */
    public static String dateToStr(Date date, String format) {
        LocalDateTime localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        return localDateTime.format(formatter);
    }

    // 時刻を持っているかどうかを判定するメソッド
    public static boolean hasTime(Date date) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);

        // 時間情報を取得
        int hours = calendar.get(Calendar.HOUR_OF_DAY);
        int minutes = calendar.get(Calendar.MINUTE);
        int seconds = calendar.get(Calendar.SECOND);

        // 時間情報があるかどうかを判定
        return hours != 0 || minutes != 0 || seconds != 0;
    }
}
