~function () {
	'use strict';
	setTimeout(DeferExecution, 1000);

}();

function DeferExecution() {
	'use strict';
	console.log("Hello from JobLeads")
	if (document.readyState=="complete") {
		
		var jobsData = GetJobs();
		//console.log(jobsData);
		//setTimeout(DeferExecution, 3000);
	} else
	{
		setTimeout(DeferExecution, 3000);
	}
	
}


function GetJobs() {
	try {
		
		var jobcards = document.querySelectorAll("html > body > table#resultsBody > tbody#resultsBodyContent > tr > td > #pageContent > tbody > tr > #resultsCol > div.jobsearch-SerpJobCard ")
		//window.AAAJobLeads = window.AAAJobLeads || {};;
		//window.AAAJobLeads.JobCards = window.AAAJobLeads.JobCards || [];
		
		var jobLeads = [];
		
		console.log(jobcards.length);
		for (var i = 0; i < jobcards.length; i++) {
			//debugger;
			var jobcardHtml = jobcards[i];
			var jobLead=ProcessJobCard(jobcardHtml,i);
			jobLeads.push(jobLead);
		}
		console.log(jobLeads);
		bytes = ConvertJobArrayTo2dMatrix(jobLeads);
		
		var xhr = new XMLHttpRequest();
		//xhr.open("POST", "http://127.0.0.1:9009/");
		xhr.open("POST", "http://127.0.0.1/");
		xhr.setRequestHeader("Content-Type", "application/octet-stream");

		xhr.onreadystatechange = function () {
		  // In local files, status is 0 upon success in Mozilla Firefox
		  if(xhr.readyState === XMLHttpRequest.DONE) {
			var status = xhr.status;
			if (status === 0 || (status >= 200 && status < 400)) {
			  // The request has been completed successfully
			  console.log("xhr.responseText:" + xhr.responseText);
			} else {
			  // Oh no! There has been an error with the request!
			  debugger;
			}
		  }
		};		

		//debugger;
		xhr.send(bytes);
		//debugger;
		
	}
	catch (err) {
		console.log("err.message: " + err.message);
		//debugger;
	}
}

function ConvertJobArrayTo2dMatrix(jobs) {
	try {
		var rows = jobs.length;
		var fieldList = ["id","jobTitle","href","company","location","salary","bullets","jobdate"]
		var columns = fieldList.length;
		var arrayConverter = new JavaScriptToVBAVariantArray();
		
		var jsGrid = arrayConverter.createGrid(rows,columns);
		//debugger;
		
		for (var jobidx = 0; jobidx < rows; jobidx++) {
			//debugger;
			var job = jobs[jobidx];
			
			for (var fieldidx =0 ; fieldidx < columns; fieldidx++) {
				
				var field = fieldList[fieldidx];
				var fieldValue = null;
				fieldValue =  job[field];
				if (fieldValue) {
					jsGrid[jobidx][fieldidx]=fieldValue;
				}
			}
		}

		var payloadEncoded = arrayConverter.persistGrid(jsGrid, rows, columns);
		
		console.log(payloadEncoded);
		return payloadEncoded;
	}
	catch (err) {
		debugger;
		console.log("ConvertJobArrayTo2dMatrix err.message: " + err.message);
		//
	}
}


function ProcessJobCard(jobcard, index) {
	try {
		//debugger;
		job = {};
		
		
		{
			job.id = jobcard.getAttribute("data-jk");
			//debugger;
			
			var titleAnchor = jobcard.querySelector("h2 > a");
			
			job.jobTitle = titleAnchor.text ;
			job.href = titleAnchor.href ;
			
		}

		{
			var sjclDiv = jobcard.querySelector("div.sjcl");
			
			try {
				var companySpan = sjclDiv.querySelector("div > span.company");
				job.company = companySpan.innerText ;
			}
			catch (err) {
				debugger;
				console.log("error whilst getting company text : " + err.message);
			}

			try {
				var locationDiv = sjclDiv.querySelector("div > div.location");
				var locationSpan = sjclDiv.querySelector("div > span.location");
				//var locationHtml = null;
				if (locationDiv) {
					job.location = locationDiv.innerText ;
				} else if (locationSpan)  {
					job.location = locationSpan.innerText ;			
				} else {
					job.location = "";
				}
			}
			catch (err) {
				debugger;
				console.log("error whilst getting location text : " + err.message);
			}
		}

		{
			try {
				//debugger;
				var salaryDiv = jobcard.querySelector("div.salarySnippet")	;
				if (salaryDiv) {
					var	salaryText = salaryDiv.querySelector("span.salary > span.salaryText");
					job.salary = salaryText.innerText;
				}
			}
			catch (err) {
				debugger;
				console.log("error whilst getting salary : " + err.message);
			}
			
		}

		{
			try {
				var jobCardShelfContainer = jobcard.querySelector("table.jobCardShelfContainer")	;
				if (jobCardShelfContainer) {
					var jobCardShelf = jobCardShelfContainer.querySelector("tbody > tr.jobCardShelf");
					if (jobCardShelf) {
						var indeedApply =  jobCardShelf.querySelector("td.indeedApply");
						if (indeedApply) {
							job.indeedApply = true;	
						}
						
						var urgentlyHiring =  jobCardShelf.querySelector("td.urgentlyHiring");
						if (urgentlyHiring) {
							job.urgentlyHiring = true;	
						}
					}
				}
			}
			catch (err) {
				debugger;
				console.log("error whilst getting summary bullet points : " + err.message);
			}
		}

		
		{
			try {
				var summaryDiv = jobcard.querySelector("div.summary")	;
				var	bulletListItems = summaryDiv.querySelectorAll("ul > li");
				var bulletPoints = [];
				for (var i = 0; i < bulletListItems.length; i++) {
					//debugger;
					var listItemText = bulletListItems[i].innerText;
					bulletPoints.push (listItemText);
				}
				
				job.bullets = bulletPoints.join('|||');
			}
			catch (err) {
				debugger;
				console.log("error whilst getting summary bullet points : " + err.message);
			}
		}

		{
			try {
				var footerDiv = jobcard.querySelector("div.jobsearch-SerpJobCard-footer")	;
				var	jobDateSpan = footerDiv.querySelector("span.date");
				
				job.jobdate = jobDateSpan.innerText;
				
			}
			catch (err) {
				debugger;
				console.log("error whilst getting job date : " + err.message);
			}
			//#p_be38c45658c4042b > div.jobsearch-SerpJobCard-footer
		}
		
		return job;
		
	}
	catch (err) {
		debugger;
		console.log("err.message: " + err.message);
	}
}


