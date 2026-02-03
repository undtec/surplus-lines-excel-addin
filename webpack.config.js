const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        functions: './src/functions/functions.js',
        taskpane: './src/taskpane/taskpane.js'
    },
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: '[name].js',
        clean: true
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                    options: {
                        presets: ['@babel/preset-env']
                    }
                }
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/functions/functions.html',
            filename: 'functions.html',
            chunks: ['functions']
        }),
        new HtmlWebpackPlugin({
            template: './src/taskpane/taskpane.html',
            filename: 'taskpane.html',
            chunks: ['taskpane'],
            inject: false
        }),
        new HtmlWebpackPlugin({
            template: './src/commands/commands.html',
            filename: 'commands.html',
            chunks: [],
            inject: false
        }),
        new CopyWebpackPlugin({
            patterns: [
                { from: 'src/functions/functions.json', to: 'functions.json' },
                { from: 'assets', to: 'assets', noErrorOnMissing: true },
                { from: 'manifest.xml', to: 'manifest.xml' }
            ]
        })
    ],
    devServer: {
        static: {
            directory: path.join(__dirname, 'dist')
        },
        headers: {
            'Access-Control-Allow-Origin': '*'
        },
        port: 3001,
        server: 'https',
        hot: true,
        allowedHosts: 'all'
    }
};
