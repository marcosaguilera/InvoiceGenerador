
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;


public class SendMail {
    
    public void SendMail() throws MessagingException {
        String host = "smtp.gmail.com";
        String Password = "Info1959+";
        String from = "info@rochester.edu.co";
        String toAddress = "maguilera@rochester.edu.co";
        String filename = "/Volumes/sapiens/INTERFASE/PSE_FILES/PSE_FILE_MAIL.txt";
        // Get system properties
        Properties props = System.getProperties();
        props.put("mail.smtp.host", host);
        props.put("mail.smtps.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        Session session = Session.getInstance(props, null);

        MimeMessage message = new MimeMessage(session);

        message.setFrom(new InternetAddress(from));
        
        message.addRecipient(
                Message.RecipientType.TO, new InternetAddress("maguilera@rochester.edu.co"));
        message.addRecipient(
                Message.RecipientType.TO, new InternetAddress("famunoz@rochester.edu.co"));
        message.addRecipient(
                Message.RecipientType.TO, new InternetAddress("efernandez@rochester.edu.co"));

        message.setSubject("Aprendoz - Archivo PSE Generado");

        BodyPart messageBodyPart = new MimeBodyPart();

        messageBodyPart.setText(
                "Estimado Usuario:"
                + "Hemos generado un nuevo archivo para PSE. Adjunto encontrará el archivo, descarguelo y carguelo a la plataforma https://www.pse.com.co ."
                + ""
                + "***Este mensaje es automatico y con fines solo informativos.***");

        Multipart multipart = new MimeMultipart();

        multipart.addBodyPart(messageBodyPart);

        messageBodyPart = new MimeBodyPart();

        DataSource source = new FileDataSource(filename);

        messageBodyPart.setDataHandler(new DataHandler(source));

        messageBodyPart.setFileName("PSE_FILE_MAIL.txt");

        multipart.addBodyPart(messageBodyPart);

        message.setContent(multipart);

        try {
            Transport tr = session.getTransport("smtps");
            tr.connect(host, from, Password);
            tr.sendMessage(message, message.getAllRecipients());
            System.out.println("Mail Sent Successfully");
            tr.close();

        } catch (SendFailedException sfe) {

            System.out.println(sfe);
        }
    }
}
