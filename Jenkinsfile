pipeline {  
	agent any
	stages { 
		//stage ('SCM Check Out') {
			//steps {
				//git url: 'https://github.com/AutomationTEst12345/ExperimentDemBobo.git'
			//}
		//}
		stage ('Test Bed Set Up') { 
			steps {
				dir('C:\\DOCKER\\MSGFW2.0'){
					bat label: '', script: 'docker build --pull --target testrunner -t msgfw2.0:test .'
				}
			}
		} 
		stage ('MSGFW2.0 Unit Test Execution') { 
			steps {
				dir('C:\\Docker\\MSGFW2.0'){
					bat label: '', script: 'docker run --rm -v "C:\\Docker\\MSGFW2.0\\Test Results":C:\\app\\msg.Testframework_cs\\TestResults msgfw2.0:test'
				}
			}
		}
	}
}
