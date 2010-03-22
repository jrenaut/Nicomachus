package docx;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import oracle.jdbc.driver.OracleDriver;
import oracle.sql.BLOB;

public class Utils {
    public static Connection openConnection() throws SQLException {
        DriverManager.registerDriver(new OracleDriver());
        return DriverManager.getConnection("jdbc:oracle:thin:@" + "//localhost:1521/epabfm", "D_BFM", "dogland");
    }

    public static BLOB createTemporaryBlob(Connection conn) throws SQLException {
        BLOB blob = BLOB.createTemporary(conn, true, BLOB.DURATION_SESSION);
        blob.open(BLOB.MODE_READWRITE);
        return blob;
    }

    public static void unzipBlobtoBlob(String fileName, BLOB input, BLOB output) throws SQLException, IOException {
        InputStream is = input.getBinaryStream();
        OutputStream out = output.setBinaryStream(1L);
        ZipInputStream in = new ZipInputStream(is);
        in.getNextEntry();
        byte[] buf = new byte[1024];
        byte[] buf2;
        int len;
        while ((len = in.read(buf)) > 0) {
            buf2 = new byte[len];
            System.arraycopy(buf, 0, buf2, 0, len);
            out.write(buf2);
        }
        out.close();
        in.close();
    }

    public static void zipBlob(String fileName, BLOB input, BLOB output) throws SQLException, IOException {
        byte[] buf = new byte[4096];
        int retval;
        InputStream is = input.binaryStreamValue();
        OutputStream os = output.setBinaryStream(1L);
        ZipOutputStream zos = new ZipOutputStream(os);
        zos.putNextEntry(new ZipEntry(fileName));
        do {
            retval = is.read(buf, 0, 1024);
            if (retval > 0) {
                zos.write(buf, 0, retval);
            }
        } while (retval != -1);
        zos.closeEntry();
        is.close();
        zos.close();
    }
}
