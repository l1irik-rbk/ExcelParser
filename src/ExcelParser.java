import javax.swing.*;
import java.awt.*;
import java.awt.event.*;


public class ExcelParser extends JFrame implements ActionListener {
    JButton loadButton, saveButton, createButton;
    JLabel selectedFileText, saveFileText;
    JFrame jf;

    public ExcelParser() {
        jf = new JFrame("Test");
        jf.setDefaultCloseOperation(EXIT_ON_CLOSE);

        jf.setLayout(new GridLayout(5, 1, 4, 4));
        jf.setVisible(true);
        setCenter();

        selectedFileText = new JLabel("Выбирете путь до файла!");
        saveFileText = new JLabel("Выбирете путь до места папки, в которую вы хотите сохранить файл!");
        loadButton = new JButton("Выбирите excel файл");
        saveButton = new JButton("Выбирете путь для сохранения excel файл");
        createButton = new JButton("Создать новый файл");

        loadButton.addActionListener(this);
        saveButton.addActionListener(this);
        createButton.addActionListener(this);

        jf.add(loadButton);
        jf.add(selectedFileText);
        jf.add(saveButton);
        jf.add(saveFileText);
        jf.add(createButton);
    }

    private void setCenter() {
        int width = 500;
        int height = 200;

        Dimension screen = Toolkit.getDefaultToolkit().getScreenSize();
        int X = (screen.width - width) / 2;
        int Y = (screen.height - height) / 2;
        jf.setBounds(X, Y, width, height);
    }
    public void actionPerformed(ActionEvent e) {

    }
}
