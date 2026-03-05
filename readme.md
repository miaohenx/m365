# MCP Server with Microsoft Graph API Integration 🚀

A Microsoft Graph API integration server built on the FastMCP framework, providing unified access interfaces for OneDrive, Mail, OneNote, and Teams functionalities.

## Features

- 🔧 **Graph API Tools Integration** - Complete Microsoft Graph API functionality support
- 📁 **OneDrive Document Management** - File upload, download, search, and management
- 📧 **Mail Services** - Email sending, receiving, and management features
- 📝 **OneNote Integration** - Notebook and page management
- 👥 **Teams Collaboration** - Teams channel and message management
- 🗄️ **MongoDB Data Storage** - Data persistence and caching
- 🐳 **Docker Containerization** - Simplified deployment and scaling
- 🔍 **Health Check** - Service status monitoring

## Quick Start

### Requirements

- Python 3.10+
- Docker & Docker Compose
- MongoDB Database
- Microsoft Graph API Access Permissions

### Installation & Deployment

1. **Clone the Project**

git clone <repository-url>
cd mcp-graph-server

2. **Configure Environment Variables**

cp .env.example .env

Edit the .env file with necessary configurations:

MONGO_URI=mongodb://localhost:27017/mcp_database
MONGO_USE_TLS=False
BASE_URL=http://localhost:8001
PORT=8001
HOST=0.0.0.0

1. **Start with Docker Compose**

docker-compose up -d

4. **Verify Service Status**

curl http://localhost:8001/health

### Manual Installation

1. **Install Dependencies**

pip install -r requirements.txt

2. **Start Service**

python server.py

## Project Structure

mcp-graph-server/
├── server.py                 # Main server file
├── tools/                    # Tools module directory
│   ├── __init__.py          # Graph API base tools
│   ├── onedrive_doc_tools.py # OneDrive document tools
│   ├── onedrive_mail_tools.py # Mail tools
│   ├── onedrive_note_tools.py # OneNote tools
│   └── onedrive_teams_tools.py # Teams tools
├── services/                 # Service layer
│   ├── onedrive_service.py  # OneDrive service
│   └── mongo_service.py     # MongoDB service
├── requirements.txt         # Python dependencies
├── Dockerfile              # Docker build file
├── docker-compose.yml      # Docker Compose configuration
└── .env                    # Environment variables configuration

## API Endpoints

### Health Check
- GET /health - Service health status check

### Graph API Tools
The server automatically registers the following tool categories:
- **Graph API Base Tools** - User information, authentication, etc.
- **OneDrive Document Tools** - File management and operations
- **Mail Tools** - Email sending and management
- **OneNote Tools** - Notebook and page operations
- **Teams Tools** - Team collaboration features

## Configuration

### Environment Variables

| Variable | Description | Default Value |
|----------|-------------|---------------|
| MONGO_URI | MongoDB connection string | - |
| MONGO_USE_TLS | Whether to use TLS connection | False |
| BASE_URL | Service base URL | http://localhost:8001 |
| PORT | Service port | 8001 |
| HOST | Service host address | 0.0.0.0 |

### Docker Configuration

The service uses Docker containerization deployment, supporting:
- Automatic restart policy
- Port mapping (8001:8001)
- Volume mounting for development debugging
- Proxy configuration support

## Development Guide

### Adding New Tools

1. Create a new tool module in the tools/ directory
2. Implement tool registration function
3. Import and register new tools in server.py

### Debug Mode

Use volume mounting for real-time development:

docker-compose up --build

Code changes will be automatically reflected in the container.

## Troubleshooting

### Common Issues

1. **Module Import Failure**
   - Check if all required files exist in the tools/ directory
   - Confirm Python path configuration is correct

2. **Service Cannot Start**
   - Check if port 8001 is occupied
   - Verify environment variable configuration is correct

3. **MongoDB Connection Failure**
   - Confirm MongoDB service is running
   - Check if MONGO_URI configuration is correct

### Log Viewing

View container logs:
docker-compose logs -f mcp-server

View real-time logs:
docker logs -f <container-id>

## License

MIT License

## Contributing

Issues and Pull Requests are welcome!

## Author

**Hengx Miao**  
Email: hengx.miao@intel.com

## Contact

For questions or support, please contact: hengx.miao@intel.com