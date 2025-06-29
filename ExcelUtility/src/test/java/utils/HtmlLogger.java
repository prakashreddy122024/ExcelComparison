package utils;

import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;

public class HtmlLogger implements AutoCloseable {
    private FileWriter writer;

    public HtmlLogger(String filePath) throws IOException {
        writer = new FileWriter(filePath, false);
        writer.write("<html><head><title>Excel Compare Log</title></head><body>\n");
        writer.write("<h2>Excel Compare Log</h2>\n");
    }

    public void log(String message) throws IOException {
        writer.write("<p>" + message + "</p>\n");
    }

    public void logWithTimestamp(String message) throws IOException {
        writer.write("<p><b>" + LocalDateTime.now() + ":</b> " + message + "</p>\n");
    }

    public void close() throws IOException {
        writer.write("</body></html>\n");
        writer.close();
    }

    public void write(String s) {
        try {
            writer.write(s);
            writer.flush();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
