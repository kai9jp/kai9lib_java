package kai9.libs;

import java.io.IOException;
import java.util.Iterator;

import javax.servlet.http.HttpServletResponse;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import lombok.Getter;
import lombok.Setter;

/**
 * Json形式のレスポンスを生成
 * API通信用の構造体
 */
@Getter
@Setter
public class JsonResponse {

    // 一般的なエラーコードの一覧
    // https://www.itmanage.co.jp/column/http-www-request-response-statuscode/

    private int return_code = 0;
    private String msg;
    private String data;
    private JSONObject jso;

    public void Add(String key, String value) throws JSONException {
        if (jso == null) {
            jso = new JSONObject();
        }
        jso.put(key, value);
    }

    public void AddArray(String key, String value) throws JSONException {
        if (jso == null) {
            jso = new JSONObject();
        }
        // valueをJSON配列として解釈するために、適切な形式で設定する
        JSONArray jsonArray = new JSONArray(value);
        jso.put(key, jsonArray);
    }

    /**
     * Json形式のレスポンスを返す
     * 
     * @param HttpServletResponse リクエストデータ
     */
    public void SetJsonResponse(HttpServletResponse res) throws IOException {
        res.setContentType("application/json");
        res.setCharacterEncoding("UTF-8");
        res.setStatus(return_code);

        JSONObject json = new JSONObject();
        try {
            json.put("return_code", return_code);

            if (msg != null) {
                json.put("msg", msg);
            }

            if (data != null) {
                // dataがJSON形式の文字列であることを確認し、そのまま設定する
                try {
                    new JSONObject(data); // dataが正しいJSON形式であることを確認するためのコード
                    json.put("data", new JSONArray(data)); // dataをJSONArrayに変換して設定
                } catch (JSONException e) {
                    json.put("data", data); // dataがJSON形式でない場合はそのまま設定
                }
            }

            if (jso != null) {
                Iterator<String> iter = jso.keys();
                while (iter.hasNext()) {
                    String key = iter.next();
                    try {
                        Object value = jso.get(key);
                        json.put(key, value);
                    } catch (JSONException e) {
                        e.printStackTrace();
                    }
                }
            }

            // レスポンスの書き込み
            res.getWriter().write(json.toString());
            // フラッシュせずにレスポンスを終了させる
            // res.getWriter().flush(); // この行を削除してフラッシュしないようにする

        } catch (JSONException e) {
            e.printStackTrace();
            // JSON生成中にエラーが発生した場合の処理を記述
        }
    }

}
