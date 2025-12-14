module.exports = {
    apps: [{
      name: 'creativ-backend',
        script: '.\\dist\\src\\main.js',
      cwd: 'C:\\Users\\Creativuk\\creativ-solar-app\\apps\\backend',
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      max_memory_restart: '512M',
      time: true,
      out_file: 'C:\\Users\\Creativuk\\creativ-solar-app\\apps\\backend\\logs\\out.log',
      error_file: 'C:\\Users\\Creativuk\\creativ-solar-app\\apps\\backend\\logs\\err.log',
      merge_logs: true,
      env: { NODE_ENV: 'development', PORT: '3000' },
      env_production: {
        NODE_ENV: 'production',
        PORT: '3000',
        API_BASE_URL: 'https://creativuk-app.paldev.tech'
      }
    }]
  };
  