package gui;

import main.Executor;
import main.Main;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class GUI implements Runnable {
    private JFrame frame;
    private static String root;
    private final String months[] = {
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December"
    };
    private static JProgressBar progressBar;
    private static JButton btnFrom, resetFrom, btnTo, execute;
    private static JLabel from_month, text2;
    private static JTextPane text;
    private static JPanel buttons, control, mainText;
    private static JScrollPane filesSelected;
    private static JComboBox monthSelector;

    public GUI(String root) {
        this.root = root;
    }
    @Override
    public void run() {
        frame = new JFrame("T-excel");
        frame.setMinimumSize(new Dimension(300, 300));
        frame.setPreferredSize(new Dimension(500, 400));
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        render(frame.getContentPane());
        frame.pack();
        frame.setVisible(true);
    }
    private void render(Container container) {
        List < File > list = new ArrayList < > ();
        final File[] to = {
                null
        };
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel files", "xl", "xls", "xlsx", "xlsm");
        // To make components style look in current os-style
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        LayoutManager layout = new BoxLayout(container, BoxLayout.Y_AXIS);
        Box boxes[] = new Box[4];
        for (int i = 0; i < 4; i++) {
            boxes[i] = Box.createHorizontalBox();
            boxes[i].createGlue();
            container.add(boxes[i]);
        }
        container.setLayout(layout);

        progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);
        btnFrom = new JButton("Load projects forecasts");
        resetFrom = new JButton("Reset projects forecasts");
        btnTo = new JButton("Load project world forecast");
        execute = new JButton("Execute copying ");
        from_month = new JLabel(" from: ");

        text = new JTextPane();
        text.setText("Please select project files");
        text.setEditable(false);
        text.setFont(new Font("Serif", Font.BOLD, 16));
        //text.setLineWrap(true);
        //text.setWrapStyleWord(true);
        text2 = new JLabel();
        text2.setFont(new Font("Serif", Font.BOLD, 14));

        LocalDateTime timePoint = LocalDateTime.now();
        int currentMonth = timePoint.getMonthValue() - 1;
        monthSelector = new JComboBox(months);
        monthSelector.setSelectedIndex(currentMonth);
        buttons = new JPanel();
        control = new JPanel();
        mainText = new JPanel();
        filesSelected = new JScrollPane(text);

        buttons.setPreferredSize(new Dimension(300, 100));
        buttons.setMinimumSize(new Dimension(300, 100));

        control.setPreferredSize(new Dimension(300, 70));
        control.setMinimumSize(new Dimension(300, 50));

        mainText.setPreferredSize(new Dimension(300, 75));
        mainText.setMinimumSize(new Dimension(300, 50));

        filesSelected.setMinimumSize(new Dimension(0, 100));
        filesSelected.setPreferredSize(new Dimension((int) Toolkit.getDefaultToolkit().getScreenSize().getWidth(), 150));

        enableControl(false);

        buttons.add(btnFrom);
        buttons.add(resetFrom);
        buttons.add(btnTo);
        control.add(execute);
        control.add(from_month);
        control.add(monthSelector);
        mainText.add(text2);

        boxes[0].add(buttons);
        boxes[1].add(control);
        boxes[2].add(mainText);
        boxes[3].add(filesSelected);

        btnFrom.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileFilter(filter);
                fileChooser.setCurrentDirectory(new File(root));
                fileChooser.setMultiSelectionEnabled(true);
                int result = fileChooser.showOpenDialog(container);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile[] = fileChooser.getSelectedFiles();
                    Arrays.stream(selectedFile).forEach(x -> {
                        if (!list.contains(x)) list.add(x);
                    });

                    String t = "Selected files: \n";
                    for (File f: list) {
                        t += f.getName() + "\n";
                    }
                    text.setText(t);
                    text.setForeground(Color.BLACK);
                    enableControl(!list.isEmpty() && to[0] != null);
                }
            }
        });

        btnTo.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileFilter(filter);
                fileChooser.setCurrentDirectory(new File(root));
                fileChooser.setMultiSelectionEnabled(false);
                int result = fileChooser.showOpenDialog(container);
                // String to;
                if (result == JFileChooser.APPROVE_OPTION) {
                    to[0] = fileChooser.getSelectedFile();
                    text2.setText("copy to: " + fileChooser.getSelectedFile().getName());
                } else {
                    to[0] = null;
                    text2.setText("");
                }
                text2.setForeground(Color.BLACK);
                enableControl(!list.isEmpty() && to[0] != null);
            }
        });

        execute.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                List < String > pathsFrom = new ArrayList < > ();
                list.forEach(x -> pathsFrom.add(x.getAbsolutePath()));
                text.setForeground(Color.BLACK);
                text2.setForeground(Color.BLACK);
                new Thread(new Runnable() {
                    public void run() {
                        Main.setLastDir(list.get(0).getParent());
                        enableControl(false);
                        enableLoad(false);
                        try {
                            new Executor().parse(pathsFrom, to[0], monthSelector.getSelectedIndex());
                            text2.setForeground(new Color(0, 100, 0));
                            text2.setText("<html>" + text2.getText() + "<br>finished</html>");
                            to[0] = null;
                            list.clear();
                            text.setText("");
                        } catch (Exception e) {
                            e.printStackTrace();
                            //text2.setForeground(Color.RED);
                            //text2.setText("<html>" + text2.getText() + "<br>finished with errors</html>");
                            String t = text.getText();
                            text.setForeground(Color.RED);
                            text.setText(t + "\nError: " + e.getMessage());
                            enableControl(true);
                        }
                        SwingUtilities.invokeLater(new Runnable() {
                            public void run() {
                                enableLoad(true);
                                // to[0] = null;
                                //resetFrom.doClick();
                            }
                        });
                    }
                }).start();
            }
        });

        resetFrom.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                list.clear();
                text.setText("");
                enableControl(false);
            }
        });
    }

    public static void enableControl(boolean enable) {
        execute.setEnabled(enable);
        monthSelector.setEnabled(enable);
    }

    public static void enableLoad(boolean enable) {
        resetFrom.setEnabled(enable);
        btnFrom.setEnabled(enable);
        resetFrom.setEnabled(enable);
        btnTo.setEnabled(enable);
        //execute.setEnabled(enable);
        //monthSelector.setEnabled(enable);
        progressBar.setVisible(!enable);
        text2.setVisible(enable);
        if (enable) {
            mainText.remove(progressBar);
        } else {
            mainText.add(progressBar);
        }
    }
}