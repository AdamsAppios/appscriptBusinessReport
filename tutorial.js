// Constructor function representing a Shape
function Shape(color) {
  this.color = color;
}

// Encapsulation: Getter and setter for color property
Shape.prototype.getColor = function() {
  return this.color;
}

Shape.prototype.setColor = function(color) {
  this.color = color;
}

// Polymorphism: Method to calculate area of a shape
Shape.prototype.calculateArea = function() {
  console.log("Cannot calculate area of an undefined shape!");
}

// Constructor function representing a Circle, inherits from Shape constructor function
function Circle(radius, color) {
  Shape.call(this, color);
  this.radius = radius;
}

// Inheritance: Circle constructor function inherits from Shape constructor function
Circle.prototype = Object.create(Shape.prototype);
Circle.prototype.constructor = Circle;

// Encapsulation: Getter and setter for radius property
Circle.prototype.getRadius = function() {
  return this.radius;
}

Circle.prototype.setRadius = function(radius) {
  this.radius = radius;
}

// Polymorphism: Method to calculate area of a circle
Circle.prototype.calculateArea = function() {
  return Math.PI * this.radius * this.radius;
}

// Constructor function representing a Rectangle, inherits from Shape constructor function
function Rectangle(width, height, color) {
  Shape.call(this, color);
  this.width = width;
  this.height = height;
}

// Inheritance: Rectangle constructor function inherits from Shape constructor function
Rectangle.prototype = Object.create(Shape.prototype);
Rectangle.prototype.constructor = Rectangle;

// Encapsulation: Getter and setter for width and height properties
Rectangle.prototype.getWidth = function() {
  return this.width;
}

Rectangle.prototype.setWidth = function(width) {
  this.width = width;
}

Rectangle.prototype.getHeight = function() {
  return this.height;
}

Rectangle.prototype.setHeight = function(height) {
  this.height = height;
}

// Polymorphism: Method to calculate area of a rectangle
Rectangle.prototype.calculateArea = function() {
  return this.width * this.height;
}

function executeOrder() {
  // Create instances of Circle and Rectangle
  var circle = new Circle(5, "red");
  var rectangle = new Rectangle(10, 5, "blue");

  // Demonstrate inheritance
  console.log(circle instanceof Shape); // true
  console.log(rectangle instanceof Shape); // true

  // Demonstrate polymorphism
  console.log(circle.calculateArea()); // 78.53981633974483
  console.log(rectangle.calculateArea()); // 50

  // Demonstrate encapsulation
  circle.setColor("green");
  rectangle.setWidth(8);

  console.log(circle.getColor()); // "green"
  console.log(rectangle.getWidth()); // 8
}