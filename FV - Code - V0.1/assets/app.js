// page load event
document.addEventListener("DOMContentLoaded", function(event) {
    var w = document.body.offsetWidth - 13, h_one_row = 30, parseDate = d3.time.format("%m/%d/%y").parse;
    var margin = {left: 400, top: 50, bottom: 50};
    var c10 = d3.scale.category10();
    var data = [];
    var sortField = null;

    var oReq = new XMLHttpRequest();
    oReq.open("GET", 'ProductionTimeLineData.xlsx', true);
    oReq.responseType = "arraybuffer";

    oReq.onload = function(e) {
        var res = new Uint8Array(oReq.response);
        var arr = new Array();
        for(var i = 0; i != res.length; ++i) arr[i] = String.fromCharCode(res[i]);
        var bstr = arr.join("");

        /* Call XLSX */
        var workbook = XLSX.read(bstr, {type:"binary"});
        var worksheet = workbook.Sheets['aaaaaaa'];
        for (z in worksheet) {
            /* all keys that do not begin with "!" correspond to cell addresses */
            if(z[0] === '!') continue;

            if(z.startsWith('A')) {
                data.push({projectId: worksheet[z].v})
            } else if(z.startsWith('B')) {
                data[data.length-1]['ProjectTitle'] = worksheet[z].v; 
            } else if(z.startsWith('C')) {
                data[data.length-1]['ProgramType'] = worksheet[z].v;
            } else if(z.startsWith('D')) {
                data[data.length-1]['TimeLineType'] = worksheet[z].v;
            } else if(z.startsWith('E')) {
                data[data.length-1]['TimeLineStartDate'] = worksheet[z].w;
            } else if(z.startsWith('F')) {
                data[data.length-1]['TimeLineEndDate'] = worksheet[z].w;
            }
        }
        // remove titles
        data.splice(0, 1);
        // filter and reverse
        data = data.filter(function(elem){
            return elem.TimeLineEndDate && elem.TimeLineStartDate;
        }).reverse();

        // create legend
        var time_types = data.map(function(elem) {return elem.TimeLineType;});
        var time_types_unique = time_types.unique();

        d3.select("#menu").selectAll("p")
            .data(time_types_unique)
            .enter()
            .append("p")
            .text(function(d) {return d + " ";})
                .append('span')
                .style('background', function(d) {return c10(d)});
        
        // fill filters
        var prog_types = data.map(function(elem) {return elem.ProgramType;});
        var unique_prog_types = prog_types.unique();

        d3.select("#project_type_filter").selectAll("p")
            .data(unique_prog_types)
            .enter()
            .append("option")
            .attr("value", function(d) {return d;})
            .attr("selected", "selected")
            .text(function(d) {return d;});

        // project_type_filter
        $('#project_type_filter').multipleSelect({
            width: '33%',
            onClick: filterChanged,
            onCheckAll: filterChanged,
            onUncheckAll: filterChanged
        });

        var proj_titles = data.map(function(elem) {return elem.ProjectTitle;});
        var unique_proj_titles = proj_titles.unique();

        d3.select("#project_title_filter").selectAll("p")
            .data(unique_proj_titles)
            .enter()
            .append("option")
            .attr("value", function(d) {return d;})
            .attr("selected", "selected")
            .text(function(d) {return d;});

        // project_title_filter
        $('#project_title_filter').multipleSelect({
            width: '33%',
            onClick: filterChanged,
            onCheckAll: filterChanged,
            onUncheckAll: filterChanged
        });

        // sort btn
        document.getElementById('project_type_sort').addEventListener('click', function() {
            sortField = 'ProgramType';
            filterChanged();
        }, false);

        document.getElementById('project_title_sort').addEventListener('click', function() {
            sortField = 'ProjectTitle';
            filterChanged();
        }, false);

        drawChart(data);
    };

    oReq.send();

    function filterChanged(view) {
        var selected_types = $('#project_type_filter').multipleSelect('getSelects');
        var selected_titles = $('#project_title_filter').multipleSelect('getSelects');

        var filtered_data = data.filter(function(elem) {
            return selected_types.indexOf(elem.ProgramType) >= 0 && selected_titles.indexOf(elem.ProjectTitle) >= 0;
        })

        if(sortField) {
            filtered_data = filtered_data.sort(function(a, b){
                return a[sortField] > b[sortField];
            }).reverse();
        }

        $('#timeline svg').remove();
        drawChart(filtered_data);
    }

    function drawChart(data) {
        flights = data.map(function(d){return parseDate(d.TimeLineStartDate)});
        var count_projects = d3.nest()
                .key(function(d) { return d.ProjectTitle; })
                .entries(data).length;

        var h = h_one_row * count_projects + margin.top + margin.bottom;

        svg = d3.select("#timeline")
            .append("svg")
            .attr("width", w)
            .attr("height", h);

        var scale = d3.time.scale()
            .domain([new Date(new Date().getFullYear(),0,1), d3.max(data, function(d){return parseDate(d.TimeLineEndDate);})])
            .range([margin.left, w]);

        var y = d3.scale.ordinal()
            .domain( data.map(function(d){return d.ProjectTitle;}))
            .rangeRoundBands([h - margin.bottom, margin.top]);

        var xaxis = d3.svg.axis().scale(scale).orient("bottom");

        var yaxis = d3.svg.axis().scale(y)
            .tickFormat(function(d) {
                return d;
            })
            .orient("right");

        var zoom = d3.behavior.zoom()
            .on("zoom", function(){
                svg.select("g").call(xaxis).selectAll("text").style("font-size", "10px");
                update_events();
            }).x(scale);

        svg.append("g")
            .attr("class", "xaxis")
            .attr("transform", "translate(0,0)")
            .call(xaxis)
            .selectAll("text")
                .style("font-size", "10px");

        function draw_events(dates) {
            var events = svg.append("g")
                .selectAll("rect.item").data(dates);

            events.enter()
                .append("rect")
                    .attr("class", "item")
                    .attr("x", function(d){return scale(parseDate(d.TimeLineStartDate));})   
                    .attr("y", function(d){ return y(d.ProjectTitle) + h_one_row/4})
                    .attr("width", function(d){return  scale(parseDate(d.TimeLineEndDate)) - scale(parseDate(d.TimeLineStartDate));})
                    .attr("ry", 3)
                    .attr("rx", 3)
                    .attr("stroke", '#383838')
                    .attr("stroke-width", 2) 
                    .attr("height", h_one_row/2)
                    .style("fill", function(d) {return c10(d.TimeLineType)});

            events.exit()
                .remove();
        }

        function update_events(){
            return svg.selectAll("rect.item")
                .attr("x", function(d){return scale(parseDate(d.TimeLineStartDate));})
                .attr("width", function(d){return scale(parseDate(d.TimeLineEndDate)) - scale(parseDate(d.TimeLineStartDate));})    
        }

        draw_events(data)

        var rect = svg.append("rect")
            .attr("x", 0)
            .attr("y", 0)
            .attr("width", w)
            .attr("height", h)
            .attr("fill-opacity", 0)
            .style('stroke-width', 5)
            .style("stroke", "#000")
            .call(zoom);

        // y axis
        var yaxis_elem = svg.append("g").attr("class", "yaxis");

        yaxis_elem.append('rect')
            .attr("width", margin.left)
            .attr('fill', 'white')
            .style('stroke-width', 2.5)
            .style("stroke", "#000")
            .attr("height", h);

        yaxis_elem.append('rect')
            .attr("width", margin.left/2)
            .attr('fill', 'white')
            .style('stroke-width', 2.5)
            .style("stroke", "#000")
            .attr("height", h);

        yaxis_elem
            .call(yaxis)
            .selectAll("text")
                .on('click', function(d) {
                    for(var i=0; i<data.length; i++) {
                        if(data[i].ProjectTitle == d) {
                            window.open('http://fv-crm-01/CRM/main.aspx?etn=opportunity&pagetype=entityrecord&id=' + data[i].projectId,'_blank');
                            return;
                        }
                    }
                })
                .style("font-size", "9px");

        // modify ticks
        svg.selectAll('.yaxis line')
            .attr('x1', 0)
            .attr('x2', w)
            .attr('y1', 15)
            .attr('y2', 15)
            .style("stroke", "#000");

        svg.select('.yaxis .domain').remove();

        // first tick
        d3.select(svg.selectAll('.yaxis g')[0].pop()).append('line')
            .attr('x1', 0)
            .attr('x2', w)
            .attr('y1', -12)
            .attr('y2', -12)
            .style("stroke", "#000");

        // types column
        var project_types = svg.append("g")
            .attr("transform", "translate(" + margin.left/2 + "," + 0 + ")")
            .attr("class", "project_types");

        project_types
            .call(yaxis)
            .selectAll("text")
                .text(function(d) {
                    for(var i=0; i<data.length; i++) {
                        if(data[i].ProjectTitle == d) {
                            return data[i].ProgramType;
                        }
                    }
                })
                .style("font-size", "9px");
    }
});


window.onscroll = function() {
    var menu_heigth = document.getElementById('menu').clientHeight;
    var top_offset = document.documentElement.scrollTop - menu_heigth;

    if(top_offset > 0) {
        d3.select('.xaxis').attr("transform", "translate(0," + top_offset + ")");
    } else {
        d3.select('.xaxis').attr("transform", "translate(0,0)");
    }
};

Array.prototype.unique = function() {
    var a = [];
    for (var i=0, l=this.length; i<l; i++)
        if (a.indexOf(this[i]) === -1)
            a.push(this[i]);
    return a;
}
