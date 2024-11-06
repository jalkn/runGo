import flask as fk
import waitress as wss
import mysql.connector as sql
import datetime as dt
import werkzeug.security as wsec
import os

app = fk.Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24)

def getDBconnection():
    conn = sql.connect(
        host='localhost',
        user='root',
        password='',
        database='arpa'
    )
    return conn

@app.route('/', methods=['GET', 'POST'])
def login():
    if fk.request.method == 'POST':
        correo = fk.request.form['correo']
        contrasena = fk.request.form['contrasena']
        
        conn = getDBconnection()
        cur = conn.cursor()
        try:
            cur.execute("SELECT contrasena, nombre FROM usuarios WHERE correo = %s", (correo,))
            usuario = cur.fetchone()
            
            if usuario and wsec.checkpassword_hash(usuario[0], contrasena):
                fk.session['logged_in'] = True
                fk.session['nombre'] = usuario[1]
                return fk.redirect(fk.url_for('index'))
            else:
                error = 'Correo o contraseña inválidos.'
                return fk.render_template('index.html', error=error)
            
        except sql.error as err:
            print(f"Error en el login: {err}")
            error = 'Error en el login. Inttenta de nuevo.'
            return fk.render_template('index.html', error=error)
        finally:
            cur.close()
            conn.close()
            
    return fk.render_template('index.html')
        
@app.route('/registrar', methods=['GET','POST'])
def registrar():
    if fk.request.method == 'POST':
        nombre = fk.request.form['nombre']
        apellido = fk.request.form['apellido']
        cedula = fk.request.form['cedula']
        correo = fk.request.form['correo']
        cargo = fk.request.form['cargo']
        contrasena = fk.request.form['contrasena']
        contrasena_hash = wsec.generate_password_hash(contrasena)

        conn = getDBconnection()
        cur = conn.cursor()

        try:
            cur.execute("INSERT INTO usuarios (nombre, apellido, cedula, correo, cargo, contrasena) VALUES (%s, %s, %s, %s, %s, %s)", (nombre, apellido, cedula, correo, cargo, contrasena_hash))
            conn.commit()
            return fk.redirect(fk.url_for('login'))
        except sql.Error as err:
            print(f"Error al registrar usuario: {err}")
            error = "Error al registrar usuario. Inténtalo de nuevo."  # Mensaje de error más genérico
            return fk.render_template('registrar.html', error=error)
        finally:
            cur.close()
            conn.close()
    return fk.render_template('registrar.html')

@app.route('/recuperar', methods=['GET', 'POST'])
def recuperar():
    if fk.request.method == 'POST':
        correo = fk.request.form['correo']
        # Aquí iría la lógica para recuperar la contraseña (ej. enviar un correo con un enlace de restablecimiento)
        # Por ahora, solo mostraremos un mensaje
        mensaje = f"Se ha enviado un correo a {correo} con instrucciones para recuperar tu contraseña."
        return fk.render_template('recuperar.html', mensaje=mensaje)  # Mostrar el mensaje al usuario.
    return fk.render_template('recuperar.html')
        
if __name__ == "__main__":
    wss.serve(app, host="0.0.0.0", port=8000)