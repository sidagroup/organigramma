<?php

class SiteController extends Controller
{
	/**
	 * Declares class-based actions.
	 */
	public function actions()
	{
		return array(
			// captcha action renders the CAPTCHA image displayed on the contact page
			'captcha'=>array(
				'class'=>'CCaptchaAction',
				'backColor'=>0xFFFFFF,
			),
			// page action renders "static" pages stored under 'protected/views/site/pages'
			// They can be accessed via: index.php?r=site/page&view=FileName
			'page'=>array(
				'class'=>'CViewAction',
			),
		);
	}

	/**
	 * This is the default 'index' action that is invoked
	 * when an action is not explicitly requested by users.
	 */
	public function actionIndex()
	{

		 $persone = Persons::model(); 
		 
		 //COORDINATORI
		 $criteriaC = new CDbCriteria();
		 $criteriaC->with = array('personsMasters.master');
		 $criteriaC->together=true;
		 $criteriaC->condition='RoleID=18 AND master.Enabled=1';		 
		 $coordinatori = $persone->findAll($criteriaC);

		 //PMA
 		 $criteriaPma = new CDbCriteria();
 		 $criteriaPma->with = array('personsMasters.master');
		 $criteriaPma->together=true;
		 $criteriaPma->condition='RoleID=11  AND master.Enabled=1';
		 $listapma = $persone->findAll($criteriaPma);
		 
		
		 // $listapms = $persone->findAll($criteriaPms);


		 $objPHPExcel = new PHPExcel();
         $sheet = $objPHPExcel->getActiveSheet()->setTitle('Simple');
         $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
         $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
         $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
         $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
         $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
         $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);

         $colonna = 0;
		 foreach ($coordinatori as $coordinatore) {
		 	$row = 1;
        	$sheet->setCellValueByColumnAndRow($colonna, $row++, $coordinatore->FirstName . ' ' . $coordinatore->LastName);

        	//CICLO I MASTER DI OGNI COORDINATORE
        	foreach ($coordinatore->personsMasters as $master) {

        		//CICLO I PMA
        		foreach ($listapma as $pma) {

        			//CICLO I MASTER DEI PMA
        			foreach ($pma->personsMasters as $pmamaster) {
        				
	        			if($pmamaster->MasterID == $master->MasterID){
	        				 //CERCO I PMS
							 $criteriaPms = new CDbCriteria();
					 		 $criteriaPms->with = array('personsMasters.master','personsCities');
							 $criteriaPms->together=true;
							 $criteriaPms->condition='RoleID=15  AND master.Enabled=1 AND personsCities.CityID=10';

	        				 $criteriaPms->addCondition('master.MasterID='. $master->MasterID . '');
		 					 $pmsTrovato = $persone->find($criteriaPms);
							 if($pmsTrovato){
		        				$sheet->setCellValueByColumnAndRow($colonna, $row++, $master->MasterID . ' - ' . $pma->FirstName . ' - ' . $pmsTrovato->FirstName);
								} else {
									$sheet->setCellValueByColumnAndRow($colonna, $row++, $master->MasterID . ' ' . $pma->FirstName );
								}			        		
			        		
	        			} 

        			}
        		}


        	}
        	
        	$colonna +=1; 
		 	
		 }

        
        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        $objWriter->save(Yii::app()->basePath . '/../files/exports/export.xlsx');

	
	}

	/**
	 * This is the action to handle external exceptions.
	 */
	public function actionError()
	{
		if($error=Yii::app()->errorHandler->error)
		{
			if(Yii::app()->request->isAjaxRequest)
				echo $error['message'];
			else
				$this->render('error', $error);
		}
	}

	/**
	 * Displays the contact page
	 */
	public function actionContact()
	{
		$model=new ContactForm;
		if(isset($_POST['ContactForm']))
		{
			$model->attributes=$_POST['ContactForm'];
			if($model->validate())
			{
				$name='=?UTF-8?B?'.base64_encode($model->name).'?=';
				$subject='=?UTF-8?B?'.base64_encode($model->subject).'?=';
				$headers="From: $name <{$model->email}>\r\n".
					"Reply-To: {$model->email}\r\n".
					"MIME-Version: 1.0\r\n".
					"Content-Type: text/plain; charset=UTF-8";

				mail(Yii::app()->params['adminEmail'],$subject,$model->body,$headers);
				Yii::app()->user->setFlash('contact','Thank you for contacting us. We will respond to you as soon as possible.');
				$this->refresh();
			}
		}
		$this->render('contact',array('model'=>$model));
	}

	/**
	 * Displays the login page
	 */
	public function actionLogin()
	{
		$model=new LoginForm;

		// if it is ajax validation request
		if(isset($_POST['ajax']) && $_POST['ajax']==='login-form')
		{
			echo CActiveForm::validate($model);
			Yii::app()->end();
		}

		// collect user input data
		if(isset($_POST['LoginForm']))
		{
			$model->attributes=$_POST['LoginForm'];
			// validate user input and redirect to the previous page if valid
			if($model->validate() && $model->login())
				$this->redirect(Yii::app()->user->returnUrl);
		}
		// display the login form
		$this->render('login',array('model'=>$model));
	}

	/**
	 * Logs out the current user and redirect to homepage.
	 */
	public function actionLogout()
	{
		Yii::app()->user->logout();
		$this->redirect(Yii::app()->homeUrl);
	}
}
