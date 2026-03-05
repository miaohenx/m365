FROM amr-registry.caas.intel.com/fiv-copilot/python:3.10-slim 

# Set environment variables for proxy
ENV HTTP_PROXY=http://child-prc.intel.com:913
ENV HTTPS_PROXY=http://child-prc.intel.com:913
ENV no_proxy=localhost,127.0.0.1,::1,intel.com,10.*.*.*,openai.azure.com

# Set working directory inside the container
WORKDIR /app

# Copy current directory contents into the container
COPY . .

# Upgrade pip and install dependencies from requirements.txt
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Expose port 8001
EXPOSE 8001

# Start the MCP server
CMD ["python", "server.py"]