package kai9.libs;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.tmatesoft.svn.core.SVNDepth;
import org.tmatesoft.svn.core.SVNException;
import org.tmatesoft.svn.core.SVNURL;
import org.tmatesoft.svn.core.internal.io.dav.DAVRepositoryFactory;
import org.tmatesoft.svn.core.internal.io.fs.FSRepositoryFactory;
import org.tmatesoft.svn.core.internal.io.svn.SVNRepositoryFactoryImpl;
import org.tmatesoft.svn.core.internal.wc.DefaultSVNOptions;
import org.tmatesoft.svn.core.wc.SVNClientManager;
import org.tmatesoft.svn.core.wc.SVNInfo;
import org.tmatesoft.svn.core.wc.SVNRevision;
import org.tmatesoft.svn.core.wc.SVNUpdateClient;
import org.tmatesoft.svn.core.wc.SVNWCClient;

public class SvnUtils {

    static {
        DAVRepositoryFactory.setup();
        SVNRepositoryFactoryImpl.setup();
        FSRepositoryFactory.setup();
    }

    private SvnUtils() {
    }

    /**
     * SVNリポジトリからローカルのフォルダへ最新の変更を更新するメソッド。
     *
     * @param svnUrl 更新したいSVNリポジトリのURL
     * @param localPath 更新を適用するローカルのパス
     * @param username SVNリポジトリへのアクセスに使用するユーザ名
     * @param password SVNリポジトリへのアクセスに使用するパスワード
     * @return 更新されたリビジョン番号
     * @throws SVNException SVN操作中に発生する可能性がある例外
     * @throws IOException
     */
    public static long updateSvnFolder(String svnUrl, String localPath, String username, String password, boolean isDelete, boolean isDir) throws SVNException, IOException {

        // ローカルの変更が保持されてしまう場合があるので、物理的にフォルダ内を毎回空にする
        if (isDelete) {
            File directory = new File(localPath);
            if (directory.isDirectory()) {
                for (File file : directory.listFiles()) {
                    FileUtils.forceDelete(file);
                }
            }
        }

        SVNClientManager clientManager = null;
        try {
            DAVRepositoryFactory.setup(); // リポジトリファクトリの初期化
            clientManager = SVNClientManager.newInstance(new DefaultSVNOptions(), username, password);

            // チェックアウトされていない場合はチェックアウトを実行
            File localDir = new File(localPath);
            boolean isWorkingCopy = false;
            if (localDir.exists()) {
                try {
                    SVNWCClient wcClient = clientManager.getWCClient();
                    SVNInfo info = wcClient.doInfo(localDir, SVNRevision.WORKING);
                    if (info != null && info.getURL() != null) {
                        isWorkingCopy = true;
                    }
                } catch (SVNException e) {
                    // ローカルディレクトリがワーキングコピーでない場合
                    isWorkingCopy = false;
                }
            }

            SVNURL url = SVNURL.parseURIEncoded(svnUrl);

            String fileName = "";
            // パスにピリオドが含まれ、かつ最後のピリオドが最後のスラッシュの後ろにあるかを確認
            if (!isDir) {
                // 最後のスラッシュの後の文字列を抽出してファイル名として取得
                fileName = svnUrl.substring(svnUrl.lastIndexOf('/') + 1);
            }

            if (!isWorkingCopy) {
                if (fileName.isEmpty()) {
                    // フォルダの場合
                    clientManager.getUpdateClient().doCheckout(
                            url,
                            localDir,
                            SVNRevision.HEAD,
                            SVNRevision.HEAD,
                            SVNDepth.INFINITY,
                            false);
                } else {
                    // ファイルの場合、上位ディレクトリをチェックアウト
                    clientManager.getUpdateClient().doCheckout(
                            url.removePathTail(),
                            localDir,
                            SVNRevision.HEAD,
                            SVNRevision.HEAD,
                            SVNDepth.EMPTY,
                            false);

                }
            }

            // SVN更新クライアントを取得
            SVNUpdateClient updateClient = clientManager.getUpdateClient();
            long revision = 0;
            if (fileName.isEmpty()) {
                // フォルダを更新
                // 指定されたローカルフォルダに最新の変更を更新
                revision = updateClient.doUpdate(
                        new File(localPath),
                        SVNRevision.HEAD, // 更新するリビジョンを指定
                        SVNDepth.INFINITY, // 更新の深さ
                        false, // サーバーから変更を受け取るかどうか
                        false // ローカルの変更を保持するかどうか
                );

            } else {
                // ファイルを更新
                revision = updateClient.doUpdate(
                        new File(localDir, fileName),
                        SVNRevision.HEAD,
                        SVNDepth.FILES,
                        false,
                        false);
            }

            return revision;
        } finally {
            if (clientManager != null) {
                clientManager.dispose();
            }
        }
    }

}