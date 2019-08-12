
$(document).ready(function(){
    //connect to the socket server.
    var socket = io.connect('http://' + document.domain + ':' + location.port + '/test');
    var msg_received = [];


    socket.on('redirect', function(data) {
        console.log("Received: " + data.url);
        window.location.href = data.url;
    });

    //receive details from serve

    socket.on('sysmessage', function(msg) {
        console.log("Received: " + msg.message);
        console.log("Received: " + msg.eta);
        console.log("Received: " + msg.progress);
        //maintain a list of ten numbers
        // if (msg_received.length >= 10){
        //     msg_received.shift()
        // }
        // msg_received.push(msg.message);
        // numbers_string = '';
        // for (var i = 0; i < msg_received.length; i++){
        //     numbers_string = numbers_string + '<p>' + msg_received[i].toString() + '</p>';
        // }
        // numbers_string = '<p>' + msg.message + '</p>';
        $('#h1').text(msg.message);
        // $('#h2').text("ETA: " + msg.eta.toString() + 's');
        $('#_progbar').text(msg.progress.toString() + '% Complete');
        $('#_progbar').attr('aria-valuenow', msg.progress).css('width', msg.progress.toString() + '%');
    });

});