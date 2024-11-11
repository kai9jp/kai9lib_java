package kai9.libs;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * TimeMeasurementクラスは、特定のタスクの累積経過時間を計測するためのユーティリティクラスです。
 * 複数のタイマーを管理し、開始と終了のアクションによって累積経過時間をコンソールに出力します。
 */
public class TimeMeasurement {
    // タイマー名とその開始時間、累積時間を保持するマップ
    private static final Map<String, Long> timers = new LinkedHashMap<>();
    private static final Map<String, Long> accumulatedTimes = new LinkedHashMap<>();

    /**
     * 累積時間をクリアします。
     */
    public static void clear() {
        accumulatedTimes.clear();
        timers.clear();
    }

    /**
     * 指定されたタイマーを開始し、開始時間を記録します。
     * 
     * @param timerName タイマーの名前
     * @param isDebug デバッグ用出力を行うかどうかのフラグ
     */
    public static void logTimeStart(String timerName, boolean isDebug) {
        startTimer(timerName, isDebug); // タイマーの開始
    }

    /**
     * 指定されたタイマーを終了し、経過時間を計算して累積時間に追加します。
     * 
     * @param timerName タイマーの名前
     * @param isDebug デバッグ用出力を行うかどうかのフラグ
     */
    public static void logTimeEnd(String timerName, boolean isDebug) {
        endTimer(timerName, isDebug); // タイマーの終了
    }

    /**
     * 指定されたタイマーを開始し、開始時間を記録します。
     * 
     * @param timerName タイマーの名前
     * @param isDebug デバッグ用出力を行うかどうかのフラグ
     */
    private static void startTimer(String timerName, boolean isDebug) {
        // 現在時刻を記録し、タイマーを開始
        timers.put(timerName, System.currentTimeMillis());
    }

    /**
     * 指定されたタイマーを終了し、経過時間を計算して累積時間に追加します。
     * 
     * @param timerName タイマーの名前
     * @param isDebug デバッグ用出力を行うかどうかのフラグ
     */
    private static void endTimer(String timerName, boolean isDebug) {
        // タイマーの開始時間を取得
        Long startTime = timers.get(timerName);
        if (startTime != null) {
            // 終了時刻を取得し、経過時間を計算
            long endTime = System.currentTimeMillis();
            long elapsedTime = endTime - startTime;

            // 累積時間を更新
            accumulatedTimes.put(timerName, accumulatedTimes.getOrDefault(timerName, 0L) + elapsedTime);

            // タイマーをマップから削除
            timers.remove(timerName);
        }
    }

    /**
     * 累積時間を出力します。
     * 
     * @param isDebug デバッグ用出力を行うかどうかのフラグ
     */
    public static void logPrint(boolean isDebug) {
        // 最大文字列長と最大時間長を初期化
        int maxLength = 0;
        int maxTimeLength = 0;

        // 累積時間マップを走査して最大長を計算
        for (Map.Entry<String, Long> entry : accumulatedTimes.entrySet()) {
            String timerName = entry.getKey();
            long timeInMs = entry.getValue();
            double timeInSeconds = timeInMs / 1000.0;
            String formattedTime = String.format("%.3f 秒", timeInSeconds);

            // フォーマットされた時間の長さをチェックして最大時間長を更新
            if (formattedTime.length() > maxTimeLength) {
                maxTimeLength = formattedTime.length();
            }

            // タイマー名の長さをチェックして最大文字列長を更新
            int length = timerName.codePointCount(0, timerName.length());
            if (length > maxLength) {
                maxLength = length;
            }
        }

        // 累積時間を整形して出力
        for (Map.Entry<String, Long> entry : accumulatedTimes.entrySet()) {
            String timerName = entry.getKey();
            long timeInMs = entry.getValue();
            double timeInSeconds = timeInMs / 1000.0;

            // 最大時間長に基づいてフォーマットされた時間を整形
            String formattedTime = String.format("%" + maxTimeLength + ".3f 秒", timeInSeconds);

            // 整形された時間とタイマー名を出力
            System.out.println(formattedTime + " : " + timerName);
        }
    }

}
