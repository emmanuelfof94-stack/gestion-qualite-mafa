module.exports = {
  apps: [
    {
      name: 'gestion-qualite',
      script: 'server.js',
      cwd: __dirname,
      watch: false,
      autorestart: true,
      restart_delay: 5000,
      max_restarts: 50,
      env: {
        NODE_ENV: 'production'
      },
      // Logs
      log_file:    './logs/combined.log',
      out_file:    './logs/out.log',
      error_file:  './logs/error.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss',
      // Redémarrage automatique si mémoire > 500 MB
      max_memory_restart: '500M'
    }
  ]
};
