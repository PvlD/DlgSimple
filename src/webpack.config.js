const path = require('path');
var webpack = require("webpack");
const CopyWebpackPlugin = require('copy-webpack-plugin')
var ExtractTextPlugin = require("extract-text-webpack-plugin");
const HtmlWebpackPlugin = require('html-webpack-plugin')
let version = require("./package.json").version

module.exports = (env) => {

    const is_production = process.argv.indexOf('production') !== -1;

  


    const WEBPACK_PLUGINS = [

        !is_production ? new CopyWebpackPlugin([
            {
                from: './src/**/*.js',
                to: './',
                
            }
        ]) : false

    ].filter(Boolean);

    console.log(" is_production:" + is_production.toString())
    console.log(" version:" + version)


    let optimization = {
        namedModules: true,
        namedChunks: true,
        splitChunks: {
            name: true,
            chunks: 'async',
            cacheGroups: {
             

                default: {
                    minChunks: 2,
                    priority: -20,
                    reuseExistingChunk: true,
                },
                vendors: {
                    name: "vendors",
                    test: /[\\/]node_modules[\\/]/,
                    chunks: 'all',
                    priority: -10
                }
            }
        },
        runtimeChunk: false
    };


    const isDevBuild = !is_production;
    return [

        {
          
            devtool: isDevBuild ? 'source-map' : "",
            entry: {
                "app": ["./src/addin/main.js"],
                "dlg": ["./src/addin/components/Dlg/main.js"],
                
            },
            output: {
                path: __dirname + "/dist",
                filename: "sd_" + version + "_[name].js",
                chunkFilename: '[id].[chunkhash].chunk.js'
            },
            resolve: {
                // Add '.ts' and '.tsx' as resolvable extensions.
                extensions: [".js", ".json"],
            },
            optimization: optimization,
            module: {
                rules: [
                    
                ]
            },
            externals: {
                //"$": "jQuery",
                "office-js": "Office"

            },
            plugins: [
                ...WEBPACK_PLUGINS,
                new webpack.DefinePlugin({
                    'process.env.NODE_ENV': isDevBuild ? '"development"' : '"production"'
                }),
                new webpack.ProvidePlugin({
                    $: "jquery",
                    jQuery: "jquery"
                }),
                new CopyWebpackPlugin([
                    {
                        from: './src/addin/Images',
                        to: './Images',
                        toType: 'dir'
                    }
                ])
                ,
              
                new HtmlWebpackPlugin({
                    
                    title: '',
                    chunks: ['app', "vendors"],
                    inject: "head",
                    filename: "index.html",
                    template: './src/addin/index.html',
                })
                ,
                new HtmlWebpackPlugin({

                    title: '',
                    
                    chunks: ['dlg', "vendors"],
                    inject: "head",
                    filename: "DlgSimple.html",
                    template: './src/addin/components/Dlg/index.html',
                })



            ]
        }];
};
