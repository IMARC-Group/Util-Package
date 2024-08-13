# Util-Package

To add util package, add the following line to your "requirements.txt"
~~~
git+https://github.com/IMARC-Group/Util-Package.git
~~~

then add the following line to use it.
~~~
import util
~~~

# How to use send_mail function

The first attempt is to send the mail via outlook if that fails then the API can be used.


~~~
import util

...

util.smtp_host = config["host"]
util.smtp_port = config["port"]
util.smtp_username = config["username"]
util.smtp_password = config["password"]

...

util.send_mail(
    subject="Subject here",
    message="Body here",
    recipients=["email 1", "email 2"],
)
~~~
