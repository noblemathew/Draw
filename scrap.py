import turtle
import random

screen = turtle.Screen()
screen.title("Dummy Drawing Program")
screen.bgcolor("skyblue")
screen.setup(800, 600)

pen = turtle.Turtle()
pen.speed(0)
pen.pensize(3)

def draw_rectangle(x, y, width, height, color):
    pen.penup()
    pen.goto(x, y)
    pen.pendown()
    pen.color("black", color)
    pen.begin_fill()
    for _ in range(2):
        pen.forward(width)
        pen.left(90)
        pen.forward(height)
        pen.left(90)
    pen.end_fill()

def draw_circle(x, y, radius, color):
    pen.penup()
    pen.goto(x, y - radius)
    pen.pendown()
    pen.color(color)
    pen.begin_fill()
    pen.circle(radius)
    pen.end_fill()

def draw_triangle(points, color):
    pen.penup()
    pen.goto(points[0])
    pen.pendown()
    pen.color("black", color)
    pen.begin_fill()
    for point in points[1:]:
        pen.goto(point)
    pen.goto(points[0])
    pen.end_fill()

# Grass
draw_rectangle(-400, -250, 800, 150, "limegreen")

# Sun
draw_circle(280, 200, 45, "yellow")

# House body
draw_rectangle(-120, -180, 240, 180, "orange")

# Roof
draw_triangle(
    [(-150, 0), (0, 130), (150, 0)],
    "red"
)

# Door
draw_rectangle(-30, -180, 60, 100, "brown")
draw_circle(15, -130, 4, "yellow")

# Windows
draw_rectangle(-95, -70, 50, 45, "lightblue")
draw_rectangle(45, -70, 50, 45, "lightblue")

# Window dividers
pen.color("white")
pen.pensize(4)

for x in [-70, 70]:
    pen.penup()
    pen.goto(x, -70)
    pen.pendown()
    pen.goto(x, -25)

    pen.penup()
    pen.goto(x - 25, -47)
    pen.pendown()
    pen.goto(x + 25, -47)

# Flowers
for x in range(-350, 351, 70):
    draw_circle(x, -180, 8, random.choice(
        ["pink", "purple", "red", "white"]
    ))
    pen.color("green")
    pen.pensize(3)
    pen.penup()
    pen.goto(x, -188)
    pen.pendown()
    pen.goto(x, -220)

# Title
pen.penup()
pen.goto(-140, 220)
pen.color("darkblue")
pen.write("MY DREAM HOUSE", font=("Arial", 22, "bold"))

pen.hideturtle()
turtle.done()
