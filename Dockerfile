FROM nginx:alpine
COPY index.html /usr/share/nginx/html/
COPY script.js /usr/share/nginx/html/
COPY style.css /usr/share/nginx/html/
EXPOSE 2000
RUN sed -i 's/listen\s*80;/listen 2000;/g' /etc/nginx/conf.d/default.conf
CMD ["nginx", "-g", "daemon off;"]
