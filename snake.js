let snake = [];
let food = { x: 0, y: 0 };
let direction = 'right';
let gameLoop;
let score = 0;
let isPaused = false;

const canvas = document.getElementById('gameCanvas');
const ctx = canvas.getContext('2d');
const gridSize = 20;
const gridWidth = canvas.width / gridSize;
const gridHeight = canvas.height / gridSize;

function initGame() {
    snake = [
        {x: 5, y: 5},
        {x: 4, y: 5},
        {x: 3, y: 5}
    ];
    direction = 'right';
    score = 0;
    generateFood();
    updateScore();
}

function generateFood() {
    food.x = Math.floor(Math.random() * gridWidth);
    food.y = Math.floor(Math.random() * gridHeight);

    // 确保食物不会生成在蛇身上
    while (checkCollision(food)) {
        food.x = Math.floor(Math.random() * gridWidth);
        food.y = Math.floor(Math.random() * gridHeight);
    }
}

function draw() {
    // 清空画布
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // 绘制蛇
    ctx.fillStyle = '#4CAF50';
    snake.forEach(segment => {
        ctx.fillRect(segment.x * gridSize, segment.y * gridSize, gridSize - 2, gridSize - 2);
    });

    // 绘制食物
    ctx.fillStyle = '#ff0000';
    ctx.fillRect(food.x * gridSize, food.y * gridSize, gridSize - 2, gridSize - 2);
}

function moveSnake() {
    if (isPaused) return;

    const head = {...snake[0]};

    switch(direction) {
        case 'up': head.y--; break;
        case 'down': head.y++; break;
        case 'left': head.x--; break;
        case 'right': head.x++; break;
    }

    // 边界穿越
    if (head.x < 0) head.x = gridWidth - 1;
    if (head.x >= gridWidth) head.x = 0;
    if (head.y < 0) head.y = gridHeight - 1;
    if (head.y >= gridHeight) head.y = 0;

    // 检查碰撞
    if (checkCollision(head)) {
        gameOver();
        return;
    }

    snake.unshift(head);

    // 检查是否吃到食物
    if (head.x === food.x && head.y === food.y) {
        score += 10;
        updateScore();
        generateFood();
    } else {
        snake.pop();
    }

    draw();
}

function checkCollision(position) {
    return snake.some(segment => segment.x === position.x && segment.y === position.y);
}

function updateScore() {
    document.getElementById('score').textContent = `分数: ${score}`;
}

function gameOver() {
    clearInterval(gameLoop);
    alert(`游戏结束！最终得分：${score}`);
}

function startGame() {
    clearInterval(gameLoop);
    initGame();
    gameLoop = setInterval(moveSnake, 150);
    isPaused = false;
}

function pauseGame() {
    isPaused = !isPaused;
}

// 键盘控制
document.addEventListener('keydown', (event) => {
    if (isPaused) return;

    switch(event.key) {
        case 'ArrowUp':
            if (direction !== 'down') direction = 'up';
            break;
        case 'ArrowDown':
            if (direction !== 'up') direction = 'down';
            break;
        case 'ArrowLeft':
            if (direction !== 'right') direction = 'left';
            break;
        case 'ArrowRight':
            if (direction !== 'left') direction = 'right';
            break;
    }
});

// 虚拟方向键控制
const directionButtons = {
    left: document.getElementById('left-btn'),
    right: document.getElementById('right-btn')
};

function handleDirectionButton(dir) {
    if (isPaused) return;
    
    switch(dir) {
        case 'left':
            if (direction !== 'right') direction = 'left';
            break;
        case 'right':
            if (direction !== 'left') direction = 'right';
            break;
    }
}

Object.entries(directionButtons).forEach(([dir, btn]) => {
    ['touchstart', 'mousedown'].forEach(eventType => {
        btn.addEventListener(eventType, (e) => {
            e.preventDefault();
            handleDirectionButton(dir);
        });
    });
});


// 初始化游戏
initGame();
draw();