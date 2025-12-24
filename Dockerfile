# Use nginx to serve static files
FROM nginx:alpine

# Copy the HTML file and Excel data to nginx html directory
COPY index.html /usr/share/nginx/html/
COPY revenue_analysis_current.xlsx /usr/share/nginx/html/

# Configure nginx to serve on port 8080 (Dokploy default)
RUN sed -i 's/listen       80;/listen       8080;/' /etc/nginx/conf.d/default.conf

# Expose port 8080
EXPOSE 8080

# Start nginx
CMD ["nginx", "-g", "daemon off;"]
