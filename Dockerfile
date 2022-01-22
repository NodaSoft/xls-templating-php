# Compile stage
FROM php:8.0.0-cli

RUN apt-get update
RUN apt-get install -y zlib1g zlib1g-dev
RUN apt-get install -y libpng-dev
RUN apt-get install -y libzip-dev

RUN docker-php-ext-install gd
RUN docker-php-ext-enable gd

RUN docker-php-ext-install zip
RUN docker-php-ext-enable zip

RUN php -r "copy('https://getcomposer.org/installer', 'composer-setup.php');"
RUN php -r "if (hash_file('sha384', 'composer-setup.php') === '906a84df04cea2aa72f40b5f787e49f22d4c2f19492ac310e8cba5b96ac8b64115ac402c8cd292b8a03482574915d1a8') { echo 'Installer verified'; } else { echo 'Installer corrupt'; unlink('composer-setup.php'); } echo PHP_EOL;"
RUN php composer-setup.php
RUN php -r "unlink('composer-setup.php');"
RUN mv composer.phar /usr/local/bin/composer
RUN chmod +x /usr/local/bin/composer
