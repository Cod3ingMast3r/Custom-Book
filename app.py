from flask import Flask, render_template, request, redirect, url_for, flash
from flask_pymongo import PyMongo
from passlib.hash import sha256_crypt

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Replace with a secure secret key
app.config['MONGO_URI'] = 'mongodb://localhost:27017/BookReplaceApp'  # Replace with your MongoDB URI
mongo = PyMongo(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/signin', methods=['GET', 'POST'])
def signin():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        # Check if the username exists in the database
        user_data = mongo.db.users.find_one({'username': username})

        if user_data:
            # Check if the entered password matches the stored hashed password
            if sha256_crypt.verify(password, user_data['password']):
                flash('Sign in successful!', 'success')
                return redirect(url_for('index'))
            else:
                flash('Incorrect password. Please try again.', 'danger')
        else:
            flash('Username not found. Please sign up first.', 'danger')

    return render_template('signin.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        # Check if the username already exists
        existing_user = mongo.db.users.find_one({'username': username})
        if existing_user:
            flash('Username already exists. Please choose another.', 'danger')
            return redirect(url_for('signup'))

        # Hash the password
        hashed_password = sha256_crypt.hash(password)

        # Insert the user into the MongoDB collection
        mongo.db.users.insert_one({'username': username, 'password': hashed_password})

        flash('You are now registered and can log in.', 'success')
        return redirect(url_for('index'))

    return render_template('signup.html')

if __name__ == '__main__':
    app.run(debug=True)
