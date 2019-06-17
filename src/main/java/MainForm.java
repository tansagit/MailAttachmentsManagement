import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import com.sun.mail.imap.IMAPFolder;

import API_Access.DriveAPIAccess;
import API_Access.SheetAPIAccess;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.mail.Address;
import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;
import javax.mail.search.FlagTerm;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JTextArea;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;

import java.awt.event.ActionListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.awt.event.ActionEvent;
import javax.swing.JScrollPane;
import javax.swing.JProgressBar;

public class MainForm extends JFrame {

	private JPanel contentPane;
	private JTextField txtEmail;
	private JPasswordField txtPassword;
	private JTextField txtSaveDir;
	private JTextField txtFolder;
	private JTextField txtSheet;
	private JTextField txtRange;
	private JFileChooser chooser;
	private JTextArea textArea;
	private JProgressBar progressBar;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					new MainForm().setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public MainForm() {
		setTitle("Mail Attachments Management");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 945, 659);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel lblEnterUser = new JLabel("Enter email & password to download attachments");
		lblEnterUser.setHorizontalAlignment(SwingConstants.RIGHT);
		lblEnterUser.setBounds(44, 23, 290, 22);
		contentPane.add(lblEnterUser);

		txtEmail = new JTextField();
		txtEmail.setBounds(368, 18, 221, 32);
		contentPane.add(txtEmail);
		txtEmail.setColumns(10);

		txtPassword = new JPasswordField();
		txtPassword.setBounds(601, 18, 204, 32);
		contentPane.add(txtPassword);

		JLabel lblSelectFolderTo = new JLabel("Save Directory");
		lblSelectFolderTo.setHorizontalAlignment(SwingConstants.RIGHT);
		lblSelectFolderTo.setBounds(54, 68, 279, 22);
		contentPane.add(lblSelectFolderTo);

		txtSaveDir = new JTextField();
		txtSaveDir.setColumns(10);
		txtSaveDir.setBounds(368, 63, 303, 32);
		txtSaveDir.setText("E:\\Attachments"); // default directory to save files
		contentPane.add(txtSaveDir);

		JButton btnBrowser = new JButton("Browser");
		btnBrowser.setBounds(683, 63, 122, 37);
		contentPane.add(btnBrowser);

		JLabel lblEnterGoogleDriver = new JLabel("Enter google driver folder to upload");
		lblEnterGoogleDriver.setHorizontalAlignment(SwingConstants.RIGHT);
		lblEnterGoogleDriver.setBounds(32, 113, 302, 22);
		contentPane.add(lblEnterGoogleDriver);

		txtFolder = new JTextField();
		txtFolder.setColumns(10);
		txtFolder.setBounds(366, 108, 439, 32);
		txtFolder.setText("1xGx9IUB9YUNhw5-w2WfGPUS0-T0F3ja-"); // default google drive folder to upload files : Resumes
		contentPane.add(txtFolder);

		JLabel lblEnterSheetId = new JLabel("Enter sheet id to update");
		lblEnterSheetId.setHorizontalAlignment(SwingConstants.RIGHT);
		lblEnterSheetId.setBounds(32, 158, 302, 22);
		contentPane.add(lblEnterSheetId);

		txtSheet = new JTextField();
		txtSheet.setColumns(10);
		txtSheet.setBounds(368, 153, 437, 32);
		txtSheet.setText("1ret6WEZyf1O7Pu_YCNDB4jAXGxgIV2KgUuZDE82kDEE"); // default google sheet to write : Resumes
																			// management
		contentPane.add(txtSheet);

		JLabel lblStartCell = new JLabel("Start cell");
		lblStartCell.setBounds(817, 158, 80, 22);
		contentPane.add(lblStartCell);

		txtRange = new JTextField();
		txtRange.setHorizontalAlignment(SwingConstants.CENTER);
		txtRange.setColumns(10);
		txtRange.setBounds(879, 153, 36, 32);
		txtRange.setText("C5"); // first cell of information table
		contentPane.add(txtRange);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(12, 193, 903, 315);
		contentPane.add(scrollPane);

		textArea = new JTextArea();
		scrollPane.setViewportView(textArea);

		final JButton btnStart = new JButton("Start Process");
		btnStart.setBounds(402, 559, 117, 40);
		contentPane.add(btnStart);

		progressBar = new JProgressBar();
		progressBar.setBounds(12, 523, 903, 23);
		progressBar.setValue(0);
		progressBar.setStringPainted(true);
		contentPane.add(progressBar);

		// -------end design-------------

		btnBrowser.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select folder to save files");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				//
				// disable the "All files" option.
				//
				chooser.setAcceptAllFileFilterUsed(false);
				//
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					File folder = chooser.getSelectedFile();
					String folderPath = folder.getPath();
					txtSaveDir.setText(folderPath);
				} else {
					System.out.println("No Selection ");
				}
			}
		});

		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (txtEmail.getText().isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls enter email!");
					txtEmail.requestFocus();
				} else if (String.valueOf(txtPassword.getPassword()).isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls enter password!");
					txtPassword.requestFocus();
				} else if (txtSaveDir.getText().isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls select save path!");
					txtSaveDir.requestFocus();
				} else if (txtFolder.getText().isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls enter google drive Id!");
					txtFolder.requestFocus();
				} else if (txtSheet.getText().isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls select google sheet Id!");
					txtSheet.requestFocus();
				} else if (txtRange.getText().isEmpty()) {
					JOptionPane.showMessageDialog(null, "Pls enter start cell!");
				} else {

					btnStart.setEnabled(false);
					textArea.setText("SCANNING........\t");
					String userName = txtEmail.getText().toString();
					String password = new String(txtPassword.getPassword());
					String saveDirectory = txtSaveDir.getText().trim();
					String driveFolderId = txtFolder.getText();
					String sheetId = txtSheet.getText();
					String sheetRange = txtRange.getText();

					// start swing worker
					WorkerDoSomething worker = new WorkerDoSomething(userName, password, saveDirectory, driveFolderId,
							sheetId, sheetRange, progressBar, textArea);
					worker.addPropertyChangeListener(new ProgressListener(progressBar));
					worker.execute();

				}

			}
		});
	}

	public class ProgressListener implements PropertyChangeListener {
		private JProgressBar progressBar;

		ProgressListener() {
		}

		ProgressListener(JProgressBar b) {
			this.progressBar = b;
			progressBar.setValue(0);
		}

		public void propertyChange(PropertyChangeEvent evt) {
			// Determine whether the property is progress type
			if ("progress".equals(evt.getPropertyName())) {
				progressBar.setValue((int) evt.getNewValue());
			}
		}
	}

	public class WorkerDoSomething extends SwingWorker<Integer, String> {
		private String userName;
		private String password;
		private String saveDirectory;
		private String driveFolderId;
		private String sheetId;
		private String sheetRange;
		private JTextArea textArea;
		private JProgressBar progressBar;

		public WorkerDoSomething(String userName, String password, String saveDirectory, String driveFolderId,
				String sheetId, String sheetRange, JProgressBar progressBar, JTextArea textArea) {
			this.userName = userName;
			this.password = password;
			this.saveDirectory = saveDirectory;
			this.driveFolderId = driveFolderId;
			this.sheetId = sheetId;
			this.sheetRange = sheetRange;
			this.progressBar = progressBar;
			this.textArea = textArea;
		}

		public void progressBarUpdate(int size) {
			int i = 0;
			int value;
			while (i <= size) {
				value = (i + 1) * 100 / size;
				progressBar.setValue(value);
				i++;
			}
		}

		public List<List<Object>> addData(String email, String Url) {
			List<Object> data1 = new ArrayList<Object>();
			data1.add(email);
			data1.add(Url);

			List<List<Object>> data = new ArrayList<List<Object>>();
			data.add(data1);
			return data;
		}

		@Override
		protected Integer doInBackground() throws Exception {
			// TODO Auto-generated method stub

			int size = 0;
			DriveAPIAccess driveAccess = new DriveAPIAccess();
			SheetAPIAccess sheetAccess = new SheetAPIAccess();

			Properties properties = new Properties();

			// server setting
			properties.put("mail.imap.host", "pop.gmail.com");
			properties.put("mail.imap.port", "995");

			// ssl setting
			properties.setProperty("mail.imap.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
			properties.setProperty("mail.imap.socketFactory.fallback", "false");
			properties.setProperty("mail.imap.socketFactory.port", String.valueOf("995"));

			Session session = Session.getDefaultInstance(properties);

			try {
				Store store = session.getStore("imaps");
				store.connect("imap.gmail.com", userName, password);

				Folder folderInbox = store.getFolder("INBOX");
				IMAPFolder ifolder = (IMAPFolder) folderInbox;
				ifolder.open(Folder.READ_WRITE);

				Message[] arrayMessages = ifolder.search(new FlagTerm(new Flags(Flags.Flag.SEEN), false)); // select
																											// unread
																											// mails
				size = arrayMessages.length; // size = number of unread mails
				if (size == 0) {
					System.out.println("No messages new!");
					publish("No message new!\n\n");
				} else {
					publish("Found " + size + " messages new\n\n");
					for (int i = 0; i < size; i++) {
						Message message = arrayMessages[i];
						Address[] fromAddress = message.getFrom();
						String fromEmail = fromAddress[0].toString();
						String subject = message.getSubject();
						String sentDate = message.getSentDate().toString();

						String contentType = message.getContentType();
						String messageContent = "";

						// store attachment file name, separated by comma
						String attachFiles = "";

						//if mail dont have attachments
						if(contentType.contains("multipart/ALTERNATIVE")) {
							publish("No attachment to download in " + fromEmail + " : " + subject + "\n\n");
							setProgress((i + 1) * 100 / size);
							this.progressBarUpdate(size);
						}
						else if (contentType.contains("multipart")) {
							Multipart multiPart = (Multipart) message.getContent();
							int numberOfParts = multiPart.getCount();
							for (int partCount = 0; partCount < numberOfParts; partCount++) {
								MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
								if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
									// this part is attachment
									String fileName = part.getFileName();
									attachFiles += fileName + ", ";
									String fileURL = saveDirectory + File.separator + fileName;

									// download attachment from mail
									part.saveFile(fileURL);
									publish("Downloaded file: " + fileName + "  from  " + fromEmail);

									// upload to google drive
									driveAccess.UploadFile(driveFolderId, fileURL, fileName);
									publish("  ======> Uploaded to drive");

									// update sheet
									String attachmentURL = driveAccess.getFileURL();
									List<List<Object>> values = this.addData(fromEmail, attachmentURL);
									sheetAccess.UpdateSheet(sheetId, sheetRange, values);
									// System.out.println("Da update sheet");
									publish("  ======> Updated to sheet\n\n");

									setProgress((i + 1) * 100 / size);
									this.progressBarUpdate(size);

								} else {
									// this part may be the message content
									messageContent = part.getContent().toString();
								}
							}

							if (attachFiles.length() > 1) {
								attachFiles = attachFiles.substring(0, attachFiles.length() - 2);
							}
						}

						else if (contentType.contains("text/plain") || contentType.contains("text/html")) {
							Object content = message.getContent();
							if (content != null) {
								messageContent = content.toString();
							}
						}

//						System.out.println("Message #" + (i + 1) + ":");
//						System.out.println("\t From: " + fromEmail);
//						System.out.println("\t Subject: " + subject);
//						System.out.println("\t Sent Date: " + sentDate);
//						System.out.println("\t Message: " + messageContent);
//						System.out.println("\t Attachments: " + attachFiles);
					}
				}

				// disconnect
				ifolder.close(false);
				store.close();

			} catch (NoSuchProviderException ex) {
				System.out.println("No provider for imaps.");
				ex.printStackTrace();
			} catch (MessagingException ex) {
				System.out.println("Could not connect to the message store, wrong email or password!");
				JOptionPane.showMessageDialog(null,
						"Could not connect to message store, please check email or password again!", "Access Denied",
						JOptionPane.ERROR_MESSAGE);
				ex.printStackTrace();
			}

			return size;

		}

		@Override
		protected void process(List<String> chunks) {
			for (String info : chunks) {
				textArea.append(info);
			}
		}

		@Override
		protected void done() {
			textArea.append("\t\t\t-------------COMPLETE------------");
			try {
				JOptionPane.showMessageDialog(null, "DONE!!!", "Attachments Processing Status",
						JOptionPane.INFORMATION_MESSAGE);
			} catch (Exception ignore) {
			}
		}

	}

}
