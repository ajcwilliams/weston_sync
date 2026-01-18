# Hetzner Dedicated Server Setup Guide

Complete guide to set up an AX41/AX52 for running SyncServer, PostgreSQL, and other services.

## 1. Order Server

1. Go to https://www.hetzner.com/dedicated-rootserver
2. Choose AX41 (€44/mo) or AX52 (€69/mo)
3. Location: Falkenstein (FSN1) - good for UK/EU
4. OS: Ubuntu 22.04 LTS
5. Wait for provisioning email (~15 min to few hours)

## 2. Initial Server Setup

SSH into your server:
```bash
ssh root@YOUR_SERVER_IP
```

Run initial setup:
```bash
# Update system
apt update && apt upgrade -y

# Set timezone
timedatectl set-timezone UTC

# Create non-root user
adduser deploy
usermod -aG sudo deploy

# Setup SSH key auth for deploy user
mkdir -p /home/deploy/.ssh
cp ~/.ssh/authorized_keys /home/deploy/.ssh/
chown -R deploy:deploy /home/deploy/.ssh
chmod 700 /home/deploy/.ssh
chmod 600 /home/deploy/.ssh/authorized_keys

# Disable root SSH login (after confirming deploy user works!)
# Edit /etc/ssh/sshd_config: PermitRootLogin no
# systemctl restart sshd
```

## 3. Firewall Setup

```bash
# Install and configure UFW
apt install ufw -y

ufw default deny incoming
ufw default allow outgoing

# SSH
ufw allow 22/tcp

# HTTP/HTTPS
ufw allow 80/tcp
ufw allow 443/tcp

# PostgreSQL (only if external access needed)
# ufw allow 5432/tcp

# Enable firewall
ufw enable
ufw status
```

## 4. Install PostgreSQL + TimescaleDB

```bash
# Add PostgreSQL repo
sh -c 'echo "deb http://apt.postgresql.org/pub/repos/apt $(lsb_release -cs)-pgdg main" > /etc/apt/sources.list.d/pgdg.list'
wget --quiet -O - https://www.postgresql.org/media/keys/ACCC4CF8.asc | apt-key add -

# Add TimescaleDB repo
echo "deb https://packagecloud.io/timescale/timescaledb/ubuntu/ $(lsb_release -cs) main" > /etc/apt/sources.list.d/timescaledb.list
wget --quiet -O - https://packagecloud.io/timescale/timescaledb/gpgkey | apt-key add -

apt update

# Install PostgreSQL 16 + TimescaleDB
apt install postgresql-16 timescaledb-2-postgresql-16 -y

# Configure TimescaleDB
timescaledb-tune --yes

# Restart PostgreSQL
systemctl restart postgresql
systemctl enable postgresql
```

### Create Database and User

```bash
sudo -u postgres psql
```

```sql
-- Create user
CREATE USER syncuser WITH PASSWORD 'YOUR_STRONG_PASSWORD';

-- Create database
CREATE DATABASE weston_sync OWNER syncuser;

-- Connect to database
\c weston_sync

-- Enable TimescaleDB extension (optional, for time-series features)
CREATE EXTENSION IF NOT EXISTS timescaledb;

-- Grant permissions
GRANT ALL PRIVILEGES ON DATABASE weston_sync TO syncuser;

\q
```

### Configure PostgreSQL for Local Access

Edit `/etc/postgresql/16/main/pg_hba.conf`:
```
# Add this line for local app access
local   weston_sync     syncuser                                md5
host    weston_sync     syncuser        127.0.0.1/32            md5
```

Restart: `systemctl restart postgresql`

## 5. Install .NET 8 Runtime

```bash
# Add Microsoft repo
wget https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

apt update
apt install aspnetcore-runtime-8.0 -y

# Verify
dotnet --list-runtimes
```

## 6. Install Nginx (Reverse Proxy)

```bash
apt install nginx -y
systemctl enable nginx
```

Create config `/etc/nginx/sites-available/syncserver`:
```nginx
server {
    listen 80;
    server_name sync.yourdomain.com;

    location / {
        proxy_pass http://localhost:5000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 86400;
    }
}
```

Enable:
```bash
ln -s /etc/nginx/sites-available/syncserver /etc/nginx/sites-enabled/
nginx -t
systemctl reload nginx
```

## 7. SSL with Let's Encrypt

```bash
apt install certbot python3-certbot-nginx -y

certbot --nginx -d sync.yourdomain.com

# Auto-renewal is configured automatically
```

## 8. Deploy SyncServer

```bash
# As deploy user
su - deploy

# Clone repo
git clone https://github.com/ajcwilliams/weston_sync.git
cd weston_sync/src/SyncServer

# Create production config
cp appsettings.json appsettings.Production.json
```

Edit `appsettings.Production.json`:
```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*",
  "Urls": "http://localhost:5000",
  "ConnectionStrings": {
    "PostgreSQL": "Host=localhost;Database=weston_sync;Username=syncuser;Password=YOUR_STRONG_PASSWORD"
  }
}
```

Build and publish:
```bash
dotnet publish -c Release -o /home/deploy/apps/syncserver
```

## 9. Create Systemd Service

Create `/etc/systemd/system/syncserver.service`:
```ini
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
```

Enable and start:
```bash
systemctl daemon-reload
systemctl enable syncserver
systemctl start syncserver
systemctl status syncserver
```

## 10. Setup Backups

Create `/home/deploy/scripts/backup.sh`:
```bash
#!/bin/bash
BACKUP_DIR="/home/deploy/backups"
DATE=$(date +%Y%m%d_%H%M%S)

mkdir -p $BACKUP_DIR

# Backup PostgreSQL
pg_dump -U syncuser weston_sync > $BACKUP_DIR/weston_sync_$DATE.sql

# Keep only last 7 days
find $BACKUP_DIR -name "*.sql" -mtime +7 -delete

echo "Backup completed: $DATE"
```

Make executable and schedule:
```bash
chmod +x /home/deploy/scripts/backup.sh

# Add to crontab (daily at 3am)
crontab -e
# Add: 0 3 * * * /home/deploy/scripts/backup.sh >> /home/deploy/logs/backup.log 2>&1
```

## 11. Verify Everything

```bash
# Check services
systemctl status postgresql
systemctl status syncserver
systemctl status nginx

# Test endpoints
curl http://localhost:5000/health
curl https://sync.yourdomain.com/health

# Check logs
journalctl -u syncserver -f
```

## Quick Reference

| Service | Port | Command |
|---------|------|---------|
| PostgreSQL | 5432 | `systemctl restart postgresql` |
| SyncServer | 5000 | `systemctl restart syncserver` |
| Nginx | 80/443 | `systemctl reload nginx` |

| Logs | Command |
|------|---------|
| SyncServer | `journalctl -u syncserver -f` |
| Nginx | `tail -f /var/log/nginx/error.log` |
| PostgreSQL | `tail -f /var/log/postgresql/postgresql-16-main.log` |

## Updating SyncServer

```bash
cd ~/weston_sync
git pull
dotnet publish -c Release -o ~/apps/syncserver src/SyncServer
sudo systemctl restart syncserver
```
