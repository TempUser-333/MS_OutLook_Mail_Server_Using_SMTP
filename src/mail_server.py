import smtplib
import ssl
import os
import sys
from flask import Flask, jsonify, request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from custom_logger import logger


app = Flask(__name__)

@app.route('/send/outlook_mail/', methods=['POST'])
def send_mail_from_outlook():
    try:
        mail = MIMEMultipart()
        
        pay_load = request.json
        
        mail['From'] = pay_load['from']
        mail['To'] = ','.join(pay_load['to']) if type(pay_load['to']) == list else pay_load['to']
        if pay_load.get('cc'):
            mail['CC'] = ','.join(pay_load['cc']) if type(pay_load['cc']) == list else pay_load['cc']
            
        if pay_load.get('bcc'):
            mail['BCC'] = ','.join(pay_load['bcc']) if type(pay_load['bcc']) == list else pay_load['bcc']
            
        if pay_load.get('body_type') == 'html':
            with open(pay_load['body'], mode='r') as f:
                try:
                    body = MIMEText(f.read(), 'html', 'utf-8')
                except Exception as ex:
                    logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
                else:
                    mail.attach(body)
                    logger.info('html body attached')
        else:
            try:
                body = MIMEText(pay_load['body'], 'plain')
            except Exception as ex:
                logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
            else:
                mail.attach(body)
                logger.info('body added')
        
        if pay_load.get('attachments'):
            if not type(pay_load['attachments']) == list:
                attachments = [pay_load['attachments']]
            else:
                attachments = pay_load['attachments']
                
            for attach in attachments:
                try:
                    with open(attach, mode='rb') as f1:
                        _attach = MIMEApplication(f1.read(), _subtype=attach.split('.')[-1])
                    _attach.add_header('Content-Disposition', 'attachment', filename=os.path.split(attach)[-1])
                except Exception as ex:
                    logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
                else:
                    mail.attach(_attach)
                    logger.info('attachment competed' + '--->' + str(attach))
        
        if pay_load.get('custom_logo'):
            try:
                with open(pay_load['custom_logo'], mode='rb') as f2:
                    image = MIMEImage(f2.read())
                image.add_header('Content-ID', '<logo>')
            except Exception as ex:
                logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
            else:
                mail.attach(image)
                logger.info('image added' + '--->' + pay_load['custom_logo'])
            
        _context = ssl.create_default_context()
        with smtplib.SMTP(pay_load['host'] if pay_load.get('host') else 'smtp.office365.com', pay_load['port'] if pay_load.get('port') else 587) as smtp:
            smtp.starttls(context=_context)
            smtp.login(pay_load['user_name'], pay_load['password'])
            logger.info(pay_load['user_name'] +  ' got logged in successfully')
            smtp.send_message(mail)
            
    except Exception as ex:
        logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
        return jsonify({'message': str((type(ex), sys.exc_info()[1], sys.exc_info()[2]))})
        
    else:
        logger.info('message sent successfully')
        return jsonify({'message': 'message sent successfully'})
        
        
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
        
