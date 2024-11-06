<!DOCTYPE html>
<html lang="es">

<head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>A R P A</title>
      <link rel="shortcut icon" href="../static/logo.png" type="image/x-icon">
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
      <link rel="stylesheet" href="../static/style.css">
</head>

<body>
  <div>
    <a href="{{ url_for('login') }}"><i class="logo" id="logo"></i></a>
  </div>

  <!-- mostrar un mensaje de error al usuario -->
  {% if error %}
      <p style="color: red;">{{ error }}</p>
  {% endif %}

<div class="container" id="container">
  <div class="form">
    <form method="post">
      <div class="input-group">
        <i class="fa fa-user"></i>
        <input type="email" id="correo" name="correo" placeholder="Usuario" required autofocus >
      </div>
      <div class="input-group">
        <i class="fa fa-lock"></i>
        <input type="password" id="contrasena" name="contrasena" placeholder="Contraseña" required>
      </div>
      <button type="submit"><i class="fa fa-sign-in"></i> Ingresar</button>
    </form>
    <div class="item">
      <a href="{{ url_for('registrar') }}"><i class="fa fa-edit"></i> Registrarse</a>
    </div>
    <div class="item">
      <a href="{{ url_for('recuperar') }}"><i class="fa fa-key"></i> Recuperar contraseña</a>
    </div>
  </div>
</div>
</body>
</html>

//recuperar
<!DOCTYPE html>
<html lang="es">

<head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Recuperar contrasena</title>
      <link rel="shortcut icon" href="../static/logo.png" type="image/x-icon">
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
      <link rel="stylesheet" href="../static/style.css">
</head>

<body>
  <div>
    <a href="{{ url_for('login') }}"><i class="logo" id="logo"></i></a>
  </div>

  <!-- mostrar un mensaje de error al usuario -->
  {% if mensaje %}
      <p>{{ mensaje }}</p>
  {% endif %}

<div class="container" id="container">
  <div class="form">
    <form method="post">
      <div class="input-group">
        <i class="fa fa-envelope"></i>
        <input type="email" id="correo" name="correo" placeholder="Correo" required>
      </div>
      <div class="input-group">
        <i class="fa fa-lock"></i>
        <input type="submit" value="Recuperar contraseña">
      </div>
    </form>
    <div class="item">
      <a href="{{ url_for('login') }}"><i class="fa fa-edit"></i>Ingresar</a>
    </div>
  </div>
</div>
</body>
</html>

//registrar
<!DOCTYPE html>
<html lang="es">

<head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Recuperar contrasena</title>
      <link rel="shortcut icon" href="../static/logo.png" type="image/x-icon">
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
      <link rel="stylesheet" href="../static/style.css">
</head>

<body>
  <div>
    <a href="{{ url_for('login') }}"><i class="logo" id="logo"></i></a>
  </div>

  <!-- mostrar un mensaje de error al usuario -->
  {% if mensaje %}
      <p>{{ mensaje }}</p>
  {% endif %}

<div class="container" id="container">
  <div class="form">
    <form method="post">
      <div class="input-group">
        <i class="fa fa-envelope"></i>
        <input type="email" id="correo" name="correo" placeholder="Correo" required>
      </div>
      <div class="input-group">
        <i class="fa fa-lock"></i>
        <input type="submit" value="Recuperar contraseña">
      </div>
    </form>
    <div class="item">
      <a href="{{ url_for('login') }}"><i class="fa fa-edit"></i>Ingresar</a>
    </div>
  </div>
</div>
</body>
</html>

//css
@import url('https://fonts.googleapis.com/css2?family=Open+Sans&display=swap');

* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}
body {
  background-color: #f5f9fa;
  font-family: 'Open Sans', sans-serif;
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 20px;
}
.logo {
  cursor: pointer;
  margin-top: 150px;
  margin-bottom: 20px;
  width: 70px;
  height: 70px;
  background-color: #3f51b5;
  position: relative;
}
.logo::before {
    content: "";
    display: block;
    width: 70px;
    height: 70px;
    background-color: #f5f9fa;
    border-radius: 50%;
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
  }
.container {
  width: 90%;
  max-width: 300px;
  overflow: hidden;
  transition: max-height 0.3s ease-out;
  max-height: 500px;
}
.form {
  text-align: left;
  padding: 20px;
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
}
.input-group {
  position: relative;
  margin-bottom: 15px;
}
.input-group i {
  position: absolute;
  left: 10px;
  top: 50%;
  transform: translateY(-50%);
  color: #3f51b5;
}
input {
  width: 100%;
  padding: 10px 10px 10px 35px;
  border: 1px solid #ddd;
  border-radius: 4px;
}
button {
  width: 100%;
  padding: 10px;
  background-color: #3f51b5;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  margin-top: 10px;
}
button:hover {
  background-color: #0056b3;
}
.item {
  margin-top: 10px;
  text-align: center;
  font-size: 12px;
  font-family: 'Open Sans', sans-serif;
}
a {
  color: #474750;
  text-decoration: none;
  display: inline-flex;
  align-items: center;
}
a:hover {
  text-decoration: underline;
}
a i {
  margin-right: 5px;
}

@media (max-width: 480px) {
  .logo {
    margin-top: 100px;
  }
  .container {
    width: 95%;
  }
  .form {
    padding: 15px;
  }
}

/*
* {
    margin: 0;
    padding: 0;
    margin-top: 40px;
    margin-bottom: -15px;
    box-sizing: border-box;
}

body {
    background-color: #50b2e3;
    color: #000;
    display: flex;
    flex-direction: column;
    align-items: center;
    font-family: 'Open Sans', sans-serif;
    font-size: 18px;
    line-height: 0.5;
}

input, button {
    font-size: 18px;
    padding: 10px 15px; 
    border-radius: 1px;
    font-family: 'Open Sans', sans-serif;
    border: none;
    appearance: none;
}

button {
    cursor: pointer;
    background-color: #00f2f6;
    transition: background-color 0.3s ease;
}

button:hover {
    background-color: #0480df;
}

.contenedor-principal {
    max-width: 900px;
    width: 90%; 
    margin: 0 auto;
}

@media (max-width: 768px) {
    .contenedor-principal {
        width: 95%;
    }
    body {
      padding: 10px;
    }
    input, button {
      padding: 8px 12px;
    }
}
*/
//estrusctura de directorios
rpapp/
├── config/
│   ├── database.php
│   └── config.php
├── controllers/
│   ├── AuthController.php
│   └── AuditController.php
├── models/
│   ├── User.php
│   └── Audit.php
├── views/
│   ├── auth/
│   │   ├── login.php
│   │   └── register.php
│   └── audits/
│       ├── index.php
│       └── create.php
└── public/
    ├── index.php
    └── .htaccess


