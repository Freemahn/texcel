package gui;

import javafx.scene.image.Image;

import javax.swing.*;

import javax.swing.filechooser.FileNameExtensionFilter;


import java.awt.Container;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class GUI implements Runnable{
    private JFrame frame;
    private static String root;

    public GUI (String root) {
        this.root = root;
    }
    @Override
    public void run() {
        frame = new JFrame("T-excel");
        frame.setPreferredSize(new Dimension(500, 500));
        frame.setIconImage(new ImageIcon(getClass().getResource("../excel.png")).getImage());
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        render(frame.getContentPane());
        frame.pack();
        frame.setVisible(true);
    }
    private void render(Container container) {
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel files", "xl", "xls", "xlsx", "xlsm");
        // To make components style look in current os-style
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch(Exception ex) {
            ex.printStackTrace();
        }

        String[] months = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
        LocalDateTime timePoint = LocalDateTime.now();
        int currentMonth = timePoint.getMonthValue() - 1;
        JComboBox monthSelector = new JComboBox(months);
        monthSelector.setSize(new Dimension(50, 50));
        monthSelector.setSelectedIndex(currentMonth);

        List<File> list = new ArrayList<>();
        final File[] to = {null};
        JButton btnFrom = new JButton("Load projects forecasts");
        JButton resetFrom = new JButton("Reset projects forecasts");
        JButton btnTo = new JButton("Load project world forecast");
        JButton execute = new JButton("Execute copying ");
        JLabel from_month = new JLabel(" from: ");
        JLabel text = new JLabel();
        JLabel text2 = new JLabel();
        Dimension d = new Dimension(50, 50);
        btnFrom.setSize(d); //setBounds(0, 0, 50, 50);
        resetFrom.setSize(d);//setBounds(0, 60, 50, 50);
        btnTo.setSize(d); //.setBounds(0, 140, 50, 50);
        execute.setSize(d); //.setBounds(0, 200, 50, 50);

        JPanel p = new JPanel();
        JPanel p2 = new JPanel();
        JPanel p3 = new JPanel();
        p.add(btnFrom);
        p.add(resetFrom);
        p.add(btnTo);
        p.add(execute);
        p.add(from_month);
        p.add(monthSelector);
        p2.add(text2);
        p3.add(text);

        p.setBounds(0, 400, 500, 200);
        p2.setBounds(0, 200, 500, 200);//.setSize(new Dimension(500, 100)); //setBounds(0, 200, 500, 100);
        p3.setSize(new Dimension(500, 200)); //.setBounds(0, 300, 500, 100);

        container.add(p);
        container.add(p2);
        container.add(p3);
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
                    Arrays.stream(selectedFile).forEach(x -> {if (!list.contains(x)) list.add(x);});

                    String t = "<html> Selected files: <br>";
                    for (File f : list) {
                        t += f.getName() + " <br> ";
                    }
                    text.setText(t + "</html>");
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

            }
        });

        execute.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (list.isEmpty()) {
                    text.setText("project forecasts not selected");
                    return;
                }
                if (to[0] == null) {
                    text2.setText("world forecast not selected");
                    return;
                }

                List<String> pathsFrom = new ArrayList<>();
                list.forEach(x -> pathsFrom.add(x.getAbsolutePath()));
                console.Main.setLastDir(list.get(0).getParent());
                int code = console.Executor.parse(pathsFrom, to[0], monthSelector.getSelectedIndex());
                text2.setText((code == 0)? "Work is done, time to get some tea" : "Some problems occured");
                to[0] = null;
                resetFrom.doClick();
            }
        });

        resetFrom.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                text.setText("");
                list.clear();
            }
        });
    }
}
