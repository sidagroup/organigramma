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

		$arrayCoordPma = array();

		 $persone = Persons::model(); 
		 
		 //COORDINATORI
		 $criteriaC = new CDbCriteria();
		 $criteriaC->with = array('personsMasters.master');
		 $criteriaC->together=true;
		 $criteriaC->condition='RoleID=18 AND master.Enabled=1';		 
		 $coordinatori = $persone->findAll($criteriaC);

		 // //PMA
 		//  $criteriaPma = new CDbCriteria();
 		//  $criteriaPma->with = array('personsMasters.master');
		 // $criteriaPma->together=true;
		 // $criteriaPma->condition='RoleID=11  AND master.Enabled=1';
		

		 // $listapma = $persone->findAll($criteriaPma);
		 
		
		 // $listapms = $persone->findAll($criteriaPms);


		 $objPHPExcel = new PHPExcel();
       
         $sheet = $objPHPExcel->getActiveSheet()->setTitle('Simple');

        

         $colonna = 10;

         $max = 0;

		 foreach ($coordinatori as $coordinatore) {
		 	$row = 4;
		 	$nPMA = 0; 


		 	//Disegno la riga dei cordinatori
        	$sheet->setCellValueByColumnAndRow($colonna, $row,  "COORDINATORE\n" . $coordinatore->FirstName ." "  . $coordinatore->LastName);

        	$row +=2; 
        	$colonnaPma = $colonna-5;
        	//CICLO I MASTER DI OGNI COORDINATORE.

        	$arrayMastersCoordinatore = array();
        	foreach ($coordinatore->personsMasters as $master) {

        		$arrayMastersCoordinatore[]=$master->MasterID;
        		//var_dump($arrayMastersCoordinatore);


        		//CICLO I PMA
        		//foreach ($listapma as $pma) {

        			//CICLO I MASTER DEI PMA
        // 			foreach ($pma->personsMasters as $pmamaster) {
        				
	       //  			if($pmamaster->MasterID == $master->MasterID){

	       //  				 $nPMA +=1;
	       //  				 //CERCO I PMS DI ANCONA
							 // $criteriaPms = new CDbCriteria();
					 		//  $criteriaPms->with = array('personsMasters.master','personsCities');
							 // $criteriaPms->together=true;
							 // $criteriaPms->condition='RoleID=15  AND master.Enabled=1 AND personsCities.CityID=10';
							 // // $criteriaPms->group="`t0_c0`";

	       //  				 $criteriaPms->addCondition('master.MasterID='. $master->MasterID . '');
		 					//  $pmsTrovato = $persone->find($criteriaPms);
							 
							 // //CERCO I PMS FUORI SEDE
							 // $criteriaPmsOutsider = new CDbCriteria();
					 		//  $criteriaPmsOutsider->with = array('personsMasters.master','personsCities');
							 // $criteriaPmsOutsider->together=true;
							 // $criteriaPmsOutsider->condition='RoleID=15  AND master.Enabled=1 AND personsCities.CityID!=10';

							 // $criteriaPmsOutsider->addCondition('master.MasterID='. $master->MasterID . '');
		 					//  $pmsTrovatoOutsider = $persone->find($criteriaPmsOutsider);

							 // if($pmsTrovato){
							 // 	$sheet->setCellValueByColumnAndRow($colonnaPma, $row, "PMA\n" . $master->master->ShortDescription . "\n\n". $pma->FirstName . ' ' . $pma->LastName );
		      //   				$sheet->setCellValueByColumnAndRow($colonnaPma++, $row+2, "PMS ANCONA \n". $pmsTrovato->FirstName . ' ' . $pmsTrovato->LastName );
								// } else {
								// 	$sheet->setCellValueByColumnAndRow($colonnaPma, $row, "PMA\n" . $master->master->ShortDescription . "\n\n". $pma->FirstName . ' ' . $pma->LastName );
		      //   				$sheet->setCellValueByColumnAndRow($colonnaPma++, $row+2, '');
								// }


			        		
	       //  			} 

        // 			}

        		//}


        	}

        	 //PMA
	 		 $criteriaPma = new CDbCriteria();
	 		 $criteriaPma->with = array('personsMasters.master');
			 $criteriaPma->together=true;
			 $criteriaPma->addCondition('RoleID=11');
			 $criteriaPma->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);

			 $listapma = $persone->findAll($criteriaPma);

			 $numeroPma = count($listapma);

			 $arrayCoordPma[$coordinatore->PersonID] = array($coordinatore,$listapma);

        	if ($max < $nPMA) $max=$nPMA;
        	
        	$colonna +=13; 
        	
		 	
		 }

		 var_dump($arrayCoordPma);


		 //var_dump($max);


		 //draw($coordinatori,);


		$styleCoordinatore = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'FFFF00'),
					),
				);

		$stylePma = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFFF'),
					),
				);

		$stylePms = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFBF'),
					),
				);


		$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(20);
        $objPHPExcel->getActiveSheet()->getRowDimension('4')->setRowHeight(50);
        $objPHPExcel->getActiveSheet()->getStyle('A4:ZZ4')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('A4:ZZ4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $objPHPExcel->getActiveSheet()->getStyle('A4:ZZ4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        
        $objPHPExcel->getActiveSheet()->getStyle('K4')->applyFromArray($styleCoordinatore);

        $objPHPExcel->getActiveSheet()->getStyle('A6:ZZ6')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('A6:ZZ6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $objPHPExcel->getActiveSheet()->getStyle('A6:ZZ6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

        $objPHPExcel->getActiveSheet()->getStyle('K6')->applyFromArray($stylePma);

        $objPHPExcel->getActiveSheet()->getStyle('A8:ZZ8')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('A8:ZZ8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $objPHPExcel->getActiveSheet()->getStyle('A8:ZZ8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

        $objPHPExcel->getActiveSheet()->getStyle('K8')->applyFromArray($stylePms);



        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        $objWriter->save(Yii::app()->basePath . '/../files/exports/export.xlsx');

	
	}


	public function draw(){


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
