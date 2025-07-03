import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;

import java.awt.*;
import java.awt.event.*;
import java.awt.print.*;
import java.io.*;
import java.text.MessageFormat;
import java.util.*;
import java.util.List;

import org.apache.logging.log4j.core.config.builder.api.Component;
import org.apache.poi.ss.usermodel.*;
import org.jfree.chart.*;
import org.jfree.chart.plot.*;
import org.jfree.data.category.*;
import org.jfree.data.general.*;
import org.jfree.data.statistics.*;

public class DataVisualization extends JFrame {
    private JTable dataTable;
    private DefaultTableModel tableModel;
    private JPanel chartPanel;
    private JComboBox<String> yColumnSelector;
    private String currentChart = "Line Plot";
    private String selectedYColumn = "";
    private JTextArea statsArea;
    private String xColumnName = "";
    private String selectedXColumn = "";
    private JComboBox<String> XColumnSelector;

    public DataVisualization() {
        setTitle("Data Visualization");
        setSize(1200, 700);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        tableModel = new DefaultTableModel();
        dataTable = new JTable(tableModel);
        add(new JScrollPane(dataTable), BorderLayout.CENTER);

        JPanel topPanel = new JPanel(new FlowLayout());

        JMenuBar menuBar = new JMenuBar();
        JMenu chartMenu = new JMenu("Select Chart");
        String[] chartTypes = {"Pie Chart", "Bar Chart", "Histogram", "Box Plot", "Line Plot"};
        for (String chart : chartTypes) {
            JMenuItem menuItem = new JMenuItem(chart);
            menuItem.addActionListener(e -> {
                currentChart = chart;
                displayChart();
            });
            chartMenu.add(menuItem);
        }
        menuBar.add(chartMenu);
        topPanel.add(menuBar);

        yColumnSelector = new JComboBox<>();
        yColumnSelector.addActionListener(e -> {
            selectedYColumn = (String) yColumnSelector.getSelectedItem();
            displayChart();
        });
        topPanel.add(new JLabel("Select Y-axis:"));
        topPanel.add(yColumnSelector);

        add(topPanel, BorderLayout.NORTH);
        XColumnSelector = new JComboBox<>();
        XColumnSelector.addActionListener(e -> {
            selectedXColumn = (String) XColumnSelector.getSelectedItem();
            displayChart();
        });
        topPanel.add(new JLabel("Select X-axis:"));
        topPanel.add(XColumnSelector);

        add(topPanel, BorderLayout.NORTH);

        JPanel rightPanel = new JPanel(new BorderLayout());
        chartPanel = new JPanel(new BorderLayout());
        chartPanel.setPreferredSize(new Dimension(500, 500));
        rightPanel.add(chartPanel, BorderLayout.CENTER);

        statsArea = new JTextArea(8, 40);
        statsArea.setEditable(false);
        JScrollPane statsScroll = new JScrollPane(statsArea);
        rightPanel.add(statsScroll, BorderLayout.NORTH);

        JButton loadFileButton = new JButton("Load Excel File");
        loadFileButton.addActionListener(this::loadExcelData);
        rightPanel.add(loadFileButton, BorderLayout.SOUTH);

        add(rightPanel, BorderLayout.EAST);
        JButton printAllButton = new JButton("Print Table + Chart");
        printAllButton.addActionListener(e -> printTableAndChart());
        
       

JPanel northPanel = new JPanel(new BorderLayout());
northPanel.add(statsScroll, BorderLayout.CENTER);
northPanel.add(printAllButton, BorderLayout.SOUTH);
rightPanel.add(northPanel, BorderLayout.NORTH);

    }
    private void printTableAndChart() {
    PrinterJob job = PrinterJob.getPrinterJob();
    job.setJobName("Print Table and Chart");

    job.setPrintable((graphics, pageFormat, pageIndex) -> {
        if (pageIndex > 0) return Printable.NO_SUCH_PAGE;

        Graphics2D g2d = (Graphics2D) graphics;
        g2d.translate(pageFormat.getImageableX(), pageFormat.getImageableY());

        // ✅ Step 1: Print the table header and contents manually
        JTableHeader header = dataTable.getTableHeader();
        Dimension headerSize = header.getPreferredSize();
        header.paint(g2d);
        g2d.translate(0, headerSize.height);
        dataTable.paint(g2d);

        // ✅ Step 2: Print the chart below the table
        g2d.translate(0, dataTable.getHeight() + 50); // spacing after table
        if (chartPanel.getComponentCount() > 0 && chartPanel.getComponent(0) instanceof ChartPanel chartComp) {
            JFreeChart chart = chartComp.getChart();
            chart.draw(g2d, new Rectangle((int) pageFormat.getImageableWidth(), 400));  // adjust height as needed
        }

        return Printable.PAGE_EXISTS;
    });

    if (job.printDialog()) {
        try {
            job.print();
            JOptionPane.showMessageDialog(this, "Table and Chart printed successfully!");
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Printing failed: " + e.getMessage());
            e.printStackTrace();
        }
    }
}



    private void loadExcelData(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(this);

        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            if (!selectedFile.getName().endsWith(".xlsx") && !selectedFile.getName().endsWith(".xls")) {
                JOptionPane.showMessageDialog(this, "Invalid Excel file format!", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            try (FileInputStream fis = new FileInputStream(selectedFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                tableModel.setRowCount(0);
                tableModel.setColumnCount(0);
                yColumnSelector.removeAllItems();
                XColumnSelector.removeAllItems();

                boolean isHeader = true;
                for (Row row : sheet) {
                    List<Object> rowData = new ArrayList<>();
                    for (int i = 0; i < row.getLastCellNum(); i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        switch (cell.getCellType()) {
                            case STRING -> rowData.add(cell.getStringCellValue());
                            case NUMERIC -> rowData.add(cell.getNumericCellValue());
                            case BOOLEAN -> rowData.add(cell.getBooleanCellValue());
                            default -> rowData.add(cell.toString());
                        }
                    }
                    if (isHeader) {
                        tableModel.setColumnIdentifiers(rowData.toArray());
                        for (int i = 0; i < rowData.size(); i++) {
                            yColumnSelector.addItem(rowData.get(i).toString());
                            XColumnSelector.addItem(rowData.get(i).toString());
                        }
                        isHeader = false;
                    } else {
                        tableModel.addRow(rowData.toArray());
                    }
                }

                if (yColumnSelector.getItemCount() > 0 )
                    yColumnSelector.setSelectedIndex(0);

                JOptionPane.showMessageDialog(this, "Excel file loaded successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                displayChart();

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(this, "Error reading Excel file: " + ex.getMessage(),
                        "File Read Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void displayChart() {
        chartPanel.removeAll();

        JFreeChart chart = switch (currentChart) {
            case "Pie Chart" -> createPieChart();
            case "Bar Chart" -> createBarChart();
            case "Histogram" -> createHistogram();
            case "Box Plot" -> createBoxPlot();
            case "Line Plot" -> createLinePlot();
            default -> null;
        };

        if (chart != null) {
            ChartPanel chartPanelComp = new ChartPanel(chart);
            chartPanel.add(chartPanelComp, BorderLayout.CENTER);
        } else {
            JOptionPane.showMessageDialog(this, "No valid chart data found!", "Chart Error", JOptionPane.ERROR_MESSAGE);
        }

        displayStats(statsArea);
        chartPanel.revalidate();
        chartPanel.repaint();
    }

    private void displayStats(JTextArea statsArea) {
        try {
            List<Double> values = new ArrayList<>();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                double val = Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex() ).toString());
                values.add(val);
            }

            double mean = values.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
            List<Double> sorted = new ArrayList<>(values);
            Collections.sort(sorted);

            double median = (sorted.size() % 2 == 0) ?
                    (sorted.get(sorted.size()/2 - 1) + sorted.get(sorted.size()/2)) / 2.0 :
                    sorted.get(sorted.size()/2);

            Map<Double, Integer> freqMap = new HashMap<>();
            for (double val : values) {
                freqMap.put(val, freqMap.getOrDefault(val, 0) + 1);
            }
            int maxFreq = Collections.max(freqMap.values());
            List<Double> modes = new ArrayList<>();
            for (var entry : freqMap.entrySet()) {
                if (entry.getValue() == maxFreq) {
                    modes.add(entry.getKey());
                }
            }

            double stdDev = Math.sqrt(values.stream().mapToDouble(v -> Math.pow(v - mean, 2)).sum() / values.size());
            double skewness = values.stream().mapToDouble(v -> Math.pow(v - mean, 3)).sum() / (values.size() * Math.pow(stdDev, 3));
            double q1 = sorted.get(sorted.size()/4);
            double q3 = sorted.get(3*sorted.size()/4);
            double iqr = q3 - q1;
            List<Double> outliers = new ArrayList<>();
            for (double val : values) {
                if (val < q1 - 1.5 * iqr || val > q3 + 1.5 * iqr)
                    outliers.add(val);
            }

            StringBuilder statsText = new StringBuilder();
            statsText.append("Mean: ").append(String.format("%.2f", mean)).append("\n");
            statsText.append("Median: ").append(String.format("%.2f", median)).append("\n");
            statsText.append("Mode: ").append(modes).append("\n");
            statsText.append("Skewness: ").append(String.format("%.2f", skewness)).append("\n");
            statsText.append("Standard Deviation: ").append(String.format("%.2f", stdDev)).append("\n");
            statsText.append("Q1: ").append(q1).append(", Q3: ").append(q3).append("\n");
            statsText.append("Outliers: ").append(outliers.isEmpty() ? "None" : outliers).append("\n");

            statsArea.setText(statsText.toString());

        } catch (Exception e) {
            statsArea.setText("Error computing statistics: " + e.getMessage());
        }
    }

    private JFreeChart createPieChart() {
        DefaultPieDataset dataset = new DefaultPieDataset();
        try {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                String category = tableModel.getValueAt(i, XColumnSelector.getSelectedIndex()).toString();

                double value = Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex() ).toString());
                dataset.setValue(category, value);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error parsing data for Pie Chart!", "Data Error", JOptionPane.ERROR_MESSAGE);
        }
        return ChartFactory.createPieChart("Pie Chart" + selectedYColumn, dataset, true, true, false);
    }

    private JFreeChart createBarChart() {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        try {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                String category = tableModel.getValueAt(i, XColumnSelector.getSelectedIndex()).toString();

                double value = Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex() ).toString());
                dataset.addValue(value, "Values", category);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error parsing data for Bar Chart!", "Data Error", JOptionPane.ERROR_MESSAGE);
        }
        return ChartFactory.createBarChart("Bar Chart of" + selectedYColumn, selectedXColumn, selectedYColumn, dataset);
    }

    private JFreeChart createHistogram() {
        HistogramDataset dataset = new HistogramDataset();
        try {
            double[] values = new double[tableModel.getRowCount()];
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                values[i] = Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex()).toString());
            }
            dataset.addSeries("Histogram", values, 10);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error parsing data for Histogram!", "Data Error", JOptionPane.ERROR_MESSAGE);
        }
        return ChartFactory.createHistogram("Histogram"+ selectedYColumn, selectedYColumn, "Frequency", dataset, PlotOrientation.VERTICAL, true, true, false);
    }

    private JFreeChart createBoxPlot() {
        DefaultBoxAndWhiskerCategoryDataset dataset = new DefaultBoxAndWhiskerCategoryDataset();
        try {
            List<Double> values = new ArrayList<>();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                values.add(Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex()).toString()));
            }
            dataset.add(values, "Series", "X-Axis: " + tableModel.getColumnName(0));
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error parsing data for Box Plot!", "Data Error", JOptionPane.ERROR_MESSAGE);
        }
        return ChartFactory.createBoxAndWhiskerChart("Box Plot", selectedYColumn, selectedYColumn, dataset, true);
    }

    private JFreeChart createLinePlot() {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        try {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                String category = tableModel.getValueAt(i, XColumnSelector.getSelectedIndex()).toString();

                double value = Double.parseDouble(tableModel.getValueAt(i, yColumnSelector.getSelectedIndex() ).toString());
                dataset.addValue(value, "Series", category);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error parsing data for Line Plot!", "Data Error", JOptionPane.ERROR_MESSAGE);
        }
        JFreeChart chart = ChartFactory.createLineChart("Line Plot", selectedXColumn, selectedYColumn, dataset, PlotOrientation.VERTICAL, true, true, false);
        CategoryPlot plot = chart.getCategoryPlot();
        plot.setDomainGridlinesVisible(true);
        plot.setRangeGridlinesVisible(true);
        return chart;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new DataVisualization().setVisible(true));
    }
}
