#!/bin/bash
# Hetzner Server Setup Script
# Run as root on a fresh Ubuntu 22.04 server

set -e

echo "=========================================="
echo "Hetzner Server Setup for Excel Sync"
echo "=========================================="

# Check if running as root
if [ "$EUID" -ne 0 ]; then
    echo "Please run as root"
    exit 1
fi

# Prompt for configuration
read -p "Enter deploy user password: " -s DEPLOY_PASS
echo
read -p "Enter PostgreSQL syncuser password: " -s PG_PASS
echo
read -p "Enter your domain (e.g., sync.example.com): " DOMAIN

echo ""
echo "Starting setup..."

# Update system
echo "[1/8] Updating system..."
apt update && apt upgrade -y

# Create deploy user
echo "[2/8] Creating deploy user..."
useradd -m -s /bin/bash deploy || true
echo "deploy:$DEPLOY_PASS" | chpasswd
usermod -aG sudo deploy

# Setup firewall
echo "[3/8] Configuring firewall..."
apt install ufw -y
ufw default deny incoming
ufw default allow outgoing
ufw allow 22/tcp
ufw allow 80/tcp
ufw allow 443/tcp
ufw --force enable

# Install PostgreSQL + TimescaleDB
echo "[4/8] Installing PostgreSQL + TimescaleDB..."
sh -c 'echo "deb http://apt.postgresql.org/pub/repos/apt $(lsb_release -cs)-pgdg main" > /etc/apt/sources.list.d/pgdg.list'
wget --quiet -O - https://www.postgresql.org/media/keys/ACCC4CF8.asc | apt-key add -
echo "deb https://packagecloud.io/timescale/timescaledb/ubuntu/ $(lsb_release -cs) main" > /etc/apt/sources.list.d/timescaledb.list
wget --quiet -O - https://packagecloud.io/timescale/timescaledb/gpgkey | apt-key add -
apt update
apt install postgresql-16 timescaledb-2-postgresql-16 -y
timescaledb-tune --yes --quiet
systemctl restart postgresql

# Create database
echo "[5/8] Creating database..."
sudo -u postgres psql -c "CREATE USER syncuser WITH PASSWORD '$PG_PASS';" || true
sudo -u postgres psql -c "CREATE DATABASE weston_sync OWNER syncuser;" || true
sudo -u postgres psql -d weston_sync -c "CREATE EXTENSION IF NOT EXISTS timescaledb;" || true

# Install .NET 8
echo "[6/8] Installing .NET 8..."
wget https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb
apt update
apt install aspnetcore-runtime-8.0 -y

# Install Nginx
echo "[7/8] Installing Nginx..."
apt install nginx -y

cat > /etc/nginx/sites-available/syncserver << EOF
server {
    listen 80;
    server_name $DOMAIN;

    location / {
        proxy_pass http://localhost:5000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_read_timeout 86400;
    }
}
EOF

ln -sf /etc/nginx/sites-available/syncserver /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx

# Create systemd service
echo "[8/8] Creating systemd service..."
cat > /etc/systemd/system/syncserver.service << EOF
[Unit]
Description=Excel Sync Server
After=network.target postgresql.service

[Service]
Type=notify
User=deploy
WorkingDirectory=/home/deploy/apps/syncserver
ExecStart=/usr/bin/dotnet /home/deploy/apps/syncserver/SyncServer.dll
Restart=always
RestartSec=10
Environment=ASPNETCORE_ENVIRONMENT=Production
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable syncserver

# Create directories
mkdir -p /home/deploy/apps /home/deploy/scripts /home/deploy/backups /home/deploy/logs
chown -R deploy:deploy /home/deploy

# Create backup script
cat > /home/deploy/scripts/backup.sh << 'EOF'
#!/bin/bash
BACKUP_DIR="/home/deploy/backups"
DATE=$(date +%Y%m%d_%H%M%S)
mkdir -p $BACKUP_DIR
PGPASSWORD=$PG_PASS pg_dump -h localhost -U syncuser weston_sync > $BACKUP_DIR/weston_sync_$DATE.sql
find $BACKUP_DIR -name "*.sql" -mtime +7 -delete
echo "Backup completed: $DATE"
EOF
chmod +x /home/deploy/scripts/backup.sh
chown deploy:deploy /home/deploy/scripts/backup.sh

echo ""
echo "=========================================="
echo "Setup complete!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "1. Point DNS for $DOMAIN to this server's IP"
echo "2. Run: certbot --nginx -d $DOMAIN"
echo "3. As deploy user, clone and deploy the app:"
echo "   su - deploy"
echo "   git clone https://github.com/ajcwilliams/weston_sync.git"
echo "   cd weston_sync/src/SyncServer"
echo "   dotnet publish -c Release -o ~/apps/syncserver"
echo ""
echo "4. Create ~/apps/syncserver/appsettings.Production.json with:"
echo "   PostgreSQL connection: Host=localhost;Database=weston_sync;Username=syncuser;Password=YOUR_PASSWORD"
echo ""
echo "5. Start the service:"
echo "   sudo systemctl start syncserver"
echo ""
echo "=========================================="
