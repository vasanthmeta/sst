package com.sstool;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.util.*;
import java.util.List;

import org.apache.logging.log4j.LogManager;

import org.apache.logging.log4j.Logger;

import org.apache.poi.util.Units;

import org.apache.poi.xwpf.usermodel.Document; import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.w3c.dom.css.Rect;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.sl.draw.geom.AdjustPointIf;

import java.awt.*;

import java.awt.event.*;

import java.awt.image.BufferedImage;

import java.io.File;

import java.io.FileOutputStream;

import java.io.IOException;

import java.io.InputStream;

import java.io.Serializable;

import java.nio.file.Files;

import java.nio.file.Path;

import java.nio.file.Paths;
import java.util.prefs.Preferences;
import java.util.stream.Stream;

@SuppressWarnings("serial")
public class WindowSnipHighlighter extends JFrame implements HighlightComponent.ClearScreenshotListener, Serializable{
private static final Logger logger = LogManager.getLogger(WindowSnipHighlighter.class);
private int counter;
private String loc;
private String path1;
private HighlightComponent highlightComponent;
private SelectiveScreenshot selectiveScreenshot;
private JTextField screenshotPath;
private JTextField documentPath;
private String home=System.getProperty("user_home") + File.separator +"Documents";
private String ssFolder = "\\Sceenshots";
private String docFolder ="\\WordDocument";
private String ssPath = home + ssFolder;
private String docPath = home + docFolder;
private transient Preferences pref;
private String ssKey ="sspath";
private String docKey = "docpath";


public WindowSnipHighlighter() {
	super("Screenshot Tool");
	setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	setSize(450,120);
	setLoc();
	
	setAlwaysOnTop(true);
	counter =1;
	screenshotPath = new JTextField();
	documentPath = new JTextField();
	JButton snipButton = new JButton("Take Snip");
	snipButton.addActionListener(e -> takeWindowSnip());
	snipButton.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_SPACE,0), "space");
	snipButton.getActionMap().put("space", new AbstractAction() {
		
		@Override
		public void actionPerformed(ActionEvent e) {
			snipButton.doClick();
		}
	});
	
	highlightComponent = new HighlightComponent();
	highlightComponent.setFocusable(true);
	highlightComponent.setClearScreenshotListener(this);
	highlightComponent.addKeyListener(new KeyAdapter() {
		@Override
		public void keyPressed(KeyEvent e) {
			if(e.getKeyCode()==KeyEvent.VK_ENTER) {
				saveScreenshot(highlightComponent.getHighlightedScreenshot());
			}else if(e.getKeyCode()==KeyEvent.VK_ESCAPE) {
				onClearScreenshot();
			}
		}
	});
	selectiveScreenshot=new SelectiveScreenshot(this);
	JButton cropButton = new JButton("Selective Snip");
	cropButton.addActionListener(e->selective());
	cropButton.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_S,0), "keyS");
	cropButton.getActionMap().put("keyS", new AbstractAction() {
		
		@Override
		public void actionPerformed(ActionEvent e) {
			cropButton.doClick();
		}
	});
	
	JButton docButton = new JButton("Create Document");
	docButton.addActionListener(e -> {
		try {
			createDocument();
		}catch(InvalidFormatException e1) {
			logger.error("Error while Creating Documnet::%s",e1.getLocalizedMessage());
		}
	});
	JButton delay = new JButton("5 Sec Delay Snip");
	delay.addActionListener(e ->{
		try {
			Thread.sleep(5000);
		}catch(InterruptedException e1) {
			Thread.currentThread().interrupt();
			e1.printStackTrace();
		}
		takeWindowSnip();
	});
	delay.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_D,0), "keyD");
	delay.getActionMap().put("keyD", new AbstractAction() {
		
		@Override
		public void actionPerformed(ActionEvent e) {
			delay.doClick();
		}
	});
	pref = Preferences.userNodeForPackage(WindowSnipHighlighter.class);
	resetPreferences();
	initializeTextFields();
	
	JButton settings = new JButton("Path Settings");
	settings.addActionListener(e -> showSettingsPopup());
	JPanel buttonPanel = new JPanel(new GridLayout(2,2));
	buttonPanel.add(snipButton);
	buttonPanel.add(cropButton);
	buttonPanel.add(delay);
	buttonPanel.add(docButton);
	buttonPanel.add(settings);
	
	add(buttonPanel,BorderLayout.NORTH);
	JScrollPane scrollPane = new JScrollPane(highlightComponent);
	scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
	scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
	add(scrollPane,BorderLayout.CENTER);
}
	private void  initializeTextFields() {
		screenshotPath.setText(pref.get(ssKey, ssPath));
		documentPath.setText(pref.get(docKey, docPath));
	}
	private void resetPreferences() {
		pref.put(ssKey, ssPath);
		pref.put(docKey, docPath);
	}
	
	private void enableSelectiveMode() {
		selectiveScreenshot.setSize(Toolkit.getDefaultToolkit().getScreenSize());
		selectiveScreenshot.setVisible(true);
		takeWindowSnip1(selectiveScreenshot.getSelectedRectangle());
	}
	
	private void selective() {
		repaint();
		enableSelectiveMode();
	}
	
	private void showSettingsPopup() {
		JDialog settDiag = new JDialog(this,"Path Settings",true);
		settDiag.setSize(300,100);
		JPanel settPanel = new JPanel();
		settPanel.setLayout(new BoxLayout(settPanel, BoxLayout.PAGE_AXIS));
		settPanel.add(new JLabel("Screenshot Path:"));
		settPanel.add(screenshotPath);
		JButton editSS = new JButton("Edit Screenshot Path:");
		settPanel.add(editSS);
		settPanel.add(new JLabel("Document PAth:"));
		settPanel.add(documentPath);
		JButton editDoc = new JButton("Edit Document Path:");
		settPanel.add(editDoc);
		
		JButton reset = new JButton("Reset");
		settPanel.add(reset);
		
		reset.addActionListener(e->{
			 screenshotPath.setText(home+ssFolder);
			 counter =1;
			 ssPath=home+ssFolder;
			 pref.put(ssKey, ssPath);
			 documentPath.setText(home+docFolder);
			 docPath=home+docFolder;
			 pref.put(docKey, docPath);
		});
		
		editSS.addActionListener(e -> handleEditButton(screenshotPath,"Screenshot"));
		editDoc.addActionListener(e -> handleEditButton(documentPath,"Document"));
		
		settDiag.add(settPanel);
		settDiag.pack();
		
		Point invokerLoc = getLocationOnScreen();
		settDiag.setLocation(invokerLoc);
		settDiag.setVisible(true);
		requestFocusInWindow();
	}
	
	private void  handleEditButton(JTextField textField, String area) {
		String currentPath = textField.getText();
		String newPath = showFileChooser(currentPath);
		
		if(newPath != null) {
		if(area.equalsIgnoreCase("Screenshot")) {
			counter =1;
			ssPath = newPath+ssFolder;
			pref.put(ssKey, ssPath);
			textField.setText(ssPath);
		}else if(area.equalsIgnoreCase("Document")) {
			docPath=newPath+docFolder;
			pref.put(docKey, docPath);
			textField.setText(docPath);
		}
		}
	}
	private String showFileChooser(String currentPath) {
		JFileChooser file = new JFileChooser(currentPath);
		file.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		
		int result=file.showOpenDialog(this);
		if(result == JFileChooser.OPEN_DIALOG) {
			File slctFile = file.getSelectedFile();
			return slctFile.getAbsolutePath();
		}
		return null;
	}
	
	public void taskBarLoc(Insets insets,DisplayMode mode) {
		String val = insets.toString();
		String[] parts = val.substring(val.indexOf('[')+1,val.indexOf(']')).split(",");
		int taskH=0;
		int x=0;
		int y=0;
		int sw=mode.getWidth();
		int sh = mode.getHeight();
		int fw=getWidth();
		int fh=getHeight();
		for (String part:parts) {
			String key = part.split("=")[0].trim();
			String value = part.split("=")[1].trim();
			if(!value.equals("0")) {
				loc = key;
			}
		}
		taskH=insets.bottom;
		x=sw-fw-50;
		y=sh-fh-taskH-50;
		
		setLocation(x,y);
	}
	
	public void setLoc() {
		GraphicsEnvironment gr=GraphicsEnvironment.getLocalGraphicsEnvironment();
		GraphicsDevice[] screens=gr.getScreenDevices();
		GraphicsDevice screen = gr.getDefaultScreenDevice();
		DisplayMode mode = screens[0].getDisplayMode();
		Insets insets = Toolkit.getDefaultToolkit().getScreenInsets(screen.getDefaultConfiguration());
		taskBarLoc(insets, mode);
	}
	
	public void onClearScreenshot() {
		setSize(450,120);
		setLoc();
	}
	
	public static void cleanUpFolder(String folderPath) {
		try {
			Path folder = Paths.get(folderPath);
			if(Files.exists(folder)) {
				try(Stream<Path> walk = Files.walk(folder)){
					walk.sorted(Comparator.reverseOrder()).map(Path::toFile).forEach(file->{
						try {
							Files.delete(file.toPath());
						}catch(IOException e) {
							e.printStackTrace();
						}
					});
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	private void takeWindowSnip1(Rectangle rect) {
		BufferedImage screenshot;
		try {
			Robot robot = new Robot();
			setExtendedState(Frame.ICONIFIED);
			screenshot = robot.createScreenCapture(rect);
			setExtendedState(Frame.NORMAL);
			highlightComponent.setScreenshot(screenshot);
			highlightComponent.clearHighlights();
			highlightComponent.repaint();
			setExtendedState(Frame.MAXIMIZED_BOTH);
			highlightComponent.requestFocusInWindow();
		}catch (AWTException ex) {
			ex.printStackTrace();
		}
	}
	private void takeWindowSnip() {
		BufferedImage screenshot;
		try {
			Robot robot = new Robot();
			setExtendedState(Frame.ICONIFIED);
			Rectangle screenBounds = getScreenBoundsExcludingTaskbar();
			screenshot = robot.createScreenCapture(screenBounds);
			setExtendedState(Frame.NORMAL);
			highlightComponent.setScreenshot(screenshot);
			highlightComponent.clearHighlights();
			highlightComponent.repaint();
			setExtendedState(Frame.MAXIMIZED_BOTH);
			highlightComponent.requestFocusInWindow();
		}catch (AWTException ex) {
			ex.printStackTrace();
		}
	}
	
	private void createDocument() throws InvalidFormatException{
		String userInput = JOptionPane.showInputDialog(this,"Enter Document Name");
		String titleInput = JOptionPane.showInputDialog(this,"Enter Title Name");
		if(userInput==null || userInput.isEmpty()) {
			JOptionPane.showMessageDialog(this, "No Input Provided.");			
		}
		String s1 = createWordDocument(createSSFolder(),userInput,titleInput);
		if(!s1.isEmpty() && s1.equalsIgnoreCase("Doc Created")) {
			int option = JOptionPane.showConfirmDialog(this, "Do you want to cleanup Screenshots Folder?","Save Screenshot",JOptionPane.YES_NO_OPTION);
			if(option == JOptionPane.YES_OPTION) {
				cleanUpFolder(createSSFolder()+File.separator);
				counter=1;
			}
			JOptionPane.showMessageDialog(this, "Documennt Created in:"+path1);
		}else {
			JOptionPane.showMessageDialog(this, "Error while creating document contct developer...");
		}
		requestFocusInWindow();
	}
	
	private Rectangle getScreenBoundsExcludingTaskbar() {
		int sw=0;
		int sh=0;
		GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
		GraphicsDevice defaultScreen = ge.getDefaultScreenDevice();
		Rectangle bounds = defaultScreen.getDefaultConfiguration().getBounds();
		Insets insets = Toolkit.getDefaultToolkit().getScreenInsets(defaultScreen.getDefaultConfiguration());
		switch(loc) {
		case "top":
			sw=bounds.width;
			sh=bounds.height-insets.top;
			return new Rectangle(bounds.x,bounds.y+insets.top,sw,sh);
		case "left":
			sw=bounds.width-insets.left;
			sh=bounds.height;
			return new Rectangle(bounds.x+insets.left,bounds.y,sw,sh);
		case "bottom":
			sw=bounds.width;
			sh=bounds.height- insets.bottom;
			return new Rectangle(bounds.x,bounds.y,sw,sh);
		case "right":
			sw=bounds.width- insets.right;
			sh=bounds.height;
			return new Rectangle(bounds.x,bounds.y,sw,sh);
		default:
			return bounds;
		}
	}
	
	private String createSSFolder() {
		String path = "";
		File newFolder = new File(ssPath);
		if(!newFolder.exists()) {
			newFolder.mkdirs();
			path = newFolder.getAbsolutePath();
		}else {
			path = newFolder.getAbsolutePath();
		}
		return path;
	}
	
	private String createdocFolder() {
		String path = "";
		File newFolder = new File(docPath);
		if(!newFolder.exists()) {
			newFolder.mkdirs();
			path = newFolder.getAbsolutePath();
		}else {
			path = newFolder.getAbsolutePath();
		}
		return path;
	
	}
	
	private void saveScreenshot(BufferedImage screenshot) {
		if(screenshot !=null) {
			try {
				File output = new File(createSSFolder()+File.separator+"image-"+counter+".png");
				ImageIO.write(screenshot, "png", output);
				JOptionPane.showMessageDialog(this, "Screenshot saved to:"+output.getAbsolutePath());
				highlightComponent.clearScreenshot();
				highlightComponent.repaint();
				counter++;
			}catch(Exception ex) {
				ex.printStackTrace();
			}
		}
	}
	
	public String createWordDocument(String imgDir,String name,String title) throws InvalidFormatException{
		String str="";
		try(XWPFDocument document = new XWPFDocument()){
			String imageDirectory = imgDir+File.separator;
			File screenshotFiles = new File(imageDirectory);
			File[] files = screenshotFiles.listFiles();
			for(int i=0;i< files.length;i++) {
				XWPFParagraph paragraph = document.createParagraph();
				String imagePath =imageDirectory+"image-"+(i+1)+".png";
				if(i==0) {
					addTextToPage(paragraph,title);
				}
				addTextToPage(paragraph,"");
				addTextToPage(paragraph,"");
				addImageToPage(paragraph,imagePath);
				addSectionBreak(document);
			}
			createWordFile(name,document);
			str="Doc Created";
		}catch(IOException ex) {
			ex.printStackTrace();
		}
		return str;
	}
	
	private void createWordFile(String name,XWPFDocument document) {
		path1=createdocFolder()+File.separator+name+".docx";
		try(FileOutputStream out = new FileOutputStream(path1)){
			document.write(out);
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void addImageToPage(XWPFParagraph paragraph,String imagePath) throws IOException,InvalidFormatException{
		if(Files.exists(Paths.get(imagePath))) {
			XWPFRun run = paragraph.createRun();
			try(InputStream input = Files.newInputStream(Paths.get(imagePath))){
				run.addPicture(input, Document.PICTURE_TYPE_PNG, imagePath, Units.toEMU(450), Units.toEMU(300));
			}catch(IOException ex) {
				ex.printStackTrace();
			}
			run.addBreak();
		}else {
			logger.error("img not found", imagePath);
		}
	}
	
	private static void addTextToPage(XWPFParagraph paragraph, String text) {
		XWPFRun run = paragraph.createRun();
		run.setText(text);
		run.addBreak();
		run.addBreak();
	}
	
	private static void addSectionBreak(XWPFDocument document) {
		XWPFParagraph paragraph = document.createParagraph();
		paragraph.setPageBreak(true);
	}
	
	public static void main(String[] args) {
		try {
			UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
		}catch(UnsupportedLookAndFeelException | ClassNotFoundException | InstantiationException |IllegalAccessException e) {
			e.printStackTrace();
		}
		SwingUtilities.invokeLater(()-> new WindowSnipHighlighter().setVisible(true));
	}
}
@SuppressWarnings("serial")
class HighlightComponent extends JComponent implements Serializable{
	private transient BufferedImage screenshot;
	private List<Rectangle> highlightAreas;
	private Rectangle currentHL;
	private transient ClearScreenshotListener clearScreenshotListener;
	
	public void setClearScreenshotListener(ClearScreenshotListener listener) {
		this.clearScreenshotListener=listener;
	}
	
	public HighlightComponent() {
		highlightAreas = new ArrayList<>();
		
		addMouseListener(new MouseAdapter() {
		@Override
		public void mousePressed(MouseEvent e) {
			currentHL=new Rectangle(e.getPoint());
			}
		@Override
		public void mouseReleased(MouseEvent e) {
			currentHL.add(e.getPoint());
			addHighlight(currentHL);
			currentHL=null;
			repaint();
		}
		});
		
		addMouseMotionListener(new MouseAdapter() {
			@Override
			public void mouseDragged(MouseEvent e) {
				if(currentHL !=null) {
					currentHL.setSize(e.getX() - currentHL.x,e.getY() - currentHL.y);
					repaint();
				}
			}
		});
	}
	
	public void setScreenshot(BufferedImage screenshot) {
		this.screenshot = screenshot;
		highlightAreas.clear();
	}
	
	public void addHighlight(Rectangle highlight) {
		highlightAreas.add(highlight);
	}
	
	public void clearHighlights() {
		highlightAreas.clear();
	}
	
	@Override
	public Dimension getPreferredSize() {
		if(screenshot !=null) {
			return new Dimension(screenshot.getWidth(),screenshot.getHeight());
		}else {
			return super.getPreferredSize();
		}
	}
	
	public BufferedImage getHighlightedScreenshot() {
		BufferedImage highlightedImage = new BufferedImage(screenshot.getWidth(),screenshot.getHeight(),BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2d = highlightedImage.createGraphics();
		g2d.drawImage(screenshot, 0, 0, this);
		for(Rectangle highlight:highlightAreas) {
			highlightedAreawithYellow(g2d,highlight);
		}
		
		g2d.dispose();
		return highlightedImage;
	}
	
	public void clearScreenshot() {
		screenshot = null;
		clearHighlights();
		if(clearScreenshotListener != null) {
			clearScreenshotListener.onClearScreenshot();
		}
	}
	
	interface ClearScreenshotListener{
		void onClearScreenshot();
	}
	
	private void highlightedAreawithYellow(Graphics2D g2d,Rectangle area) {
		Color originalColor = g2d.getColor();
		g2d.setColor(new Color(255,255,0,150));
		g2d.fillRect(area.x, area.y, area.width, area.height);
		g2d.setColor(originalColor);
	}
	
	@Override
	protected void paintComponent(Graphics g) {
		super.paintComponent(g);
		
		if(screenshot != null) {
			g.drawImage(screenshot,0,0,this);
			
			Graphics2D g2d = (Graphics2D)g;
			g2d.setColor(Color.RED);
			g2d.setStroke(new BasicStroke(2));
			if(currentHL != null) {
				g2d.drawRect(currentHL.x,currentHL.y,currentHL.width,currentHL.height);
			}
			for(Rectangle highlight:highlightAreas) {
				g2d.drawRect(highlight.x,highlight.y,highlight.width,highlight.height);
			}
		}
	}
}
@SuppressWarnings("serial")
class SelectiveScreenshot extends JDialog{
	private boolean isselectiveMode = false;
	private Point startPoint;
	private Point endPoint;
	private Rectangle selectedRectangle;
	
	
	public SelectiveScreenshot(JFrame parent) {
		super(parent,true);
		setUndecorated(true);
		setOpacity(0.1f);
		addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(MouseEvent e) {
				startPoint = e.getPoint();
			}
			@Override
			public void mouseReleased(MouseEvent e) {
				endPoint = e.getPoint();
				isselectiveMode = false;
				handleSelectedRectangle();
				dispose();
			}
		});
		
		addMouseMotionListener(new MouseAdapter() {
			@Override
			public void mouseDragged(MouseEvent e) {
				if(!isselectiveMode) {
					isselectiveMode=true;
				}
				endPoint = e.getPoint();
				repaint();
			}
		});
		addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if(e.getKeyCode()==KeyEvent.VK_ESCAPE) {
					isselectiveMode=false;
					dispose();
				}
			}
		});
		setFocusable(true);
		requestFocus();
	}
	
	private void handleSelectedRectangle() {
		int x = Math.min(startPoint.x,endPoint.x);
		int y = Math.min(startPoint.y,endPoint.y);
		int width = Math.abs(startPoint.x-endPoint.x);
		int height = Math.abs(startPoint.y-endPoint.y);
		selectedRectangle=new Rectangle(x,y,width,height);
	}
	
	public Rectangle getSelectedRectangle() {
		return selectedRectangle;
	}
	
	@Override
	public void paint(Graphics g) {
		super.paint(g);
		
		if(isselectiveMode && startPoint !=null && endPoint !=null) {
			Graphics2D g2d = (Graphics2D)g;
			g2d.setColor(new Color(0,0,255,100));
			int x = Math.min(startPoint.x,endPoint.x);
			int y = Math.min(startPoint.y,endPoint.y);
			int width = Math.abs(startPoint.x-endPoint.x);
			int height = Math.abs(startPoint.y-endPoint.y);
			g2d.fill(new Rectangle(x,y,width,height));
		}
	}
}