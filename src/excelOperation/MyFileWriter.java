package excelOperation;
import java.io.File;
import java.io.FileDescriptor;
import java.io.FileWriter;
import java.io.IOException;


public class MyFileWriter extends FileWriter {

    public MyFileWriter(String fileName) throws IOException {
        super(fileName);
        // TODO Auto-generated constructor stub
    }

    public MyFileWriter(File file) throws IOException {
        super(file);
        // TODO Auto-generated constructor stub
    }

    public MyFileWriter(FileDescriptor fd) {
        super(fd);
        // TODO Auto-generated constructor stub
    }

    public MyFileWriter(String fileName, boolean append) throws IOException {
        super(fileName, append);
        // TODO Auto-generated constructor stub
    }

    public MyFileWriter(File file, boolean append) throws IOException {
        super(file, append);
        // TODO Auto-generated constructor stub
    }
    
    public void appendLine(CharSequence c){
        try {
            super.append(c);
            super.append("\n");
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
